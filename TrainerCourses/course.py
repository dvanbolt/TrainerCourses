from __future__ import annotations
from .openpyxl_extension import open as open_xlsx,Workbook
from typing import Optional,Any,Generator,Callable,Iterator,Iterable
from enum import Enum
from copy import deepcopy,copy
from math import sqrt
from datetime import datetime
from pathlib import Path
from dataclasses import dataclass,asdict
from pydantic import BaseModel,validator,Extra
from .srch import qkfltr


FIT_VERSION = 2

#stats.mean is slow af for this use case
def mean(a:list[float])->float:
    s = 0.0
    for i in a:
        s+=i
    return s / len(a)

def pathsafe(fp:str|Path|Callable)->Callable|Path|str:
    def _clean_str(s:str):
        return "".join([c for c in s if c.isalpha() or c.isdigit() or c in (' ',"_",'-','\\',':')]).rstrip()
    
    def _clean_path(s:Path|str)->str|Path:
        if isinstance(s,str):
            return _clean_str(s)
        elif isinstance(s,Path):
            if s.absolute():
                return Path(*['C:\\']+[_clean_str(p.stem) for p in list(reversed(s.parents))+[s]])
            else:
                return Path(*[_clean_str(p.stem) for p in list(reversed(s.parents))+[s]])
    if callable(fp):
        def _inner(s):
            return _clean_path(fp(s))
        return _inner
    else:
        return _clean_path(fp)


class Templated:
    _key_word_argument = None
    class_init_kwargs = {'extra':Extra.forbid}

    @classmethod
    def remap(cls,**kwargs):
        d = {}
        for key,value in cls._key.items():
            d[value] = kwargs[key]
        return d

    @classmethod
    def parse(cls,**kwargs):            
        return cls(**cls.remap(**kwargs))

class Category(str,Enum):
    anaerobic = "Anaerobic"
    warmup = "Warmup"
    cooldown = "Cooldown"
    steadystead = "Steady State"
    sweetspot = "Sweet Spot"
    aerobic = "Aerobic Base"

class CourseCollection:
    class UserProfile(BaseModel,Templated,**Templated.class_init_kwargs):
        _key = {"Functional Threshold Power":"ftp"}
        ftp:float

    def __init__(self,name:str,user:UserProfile,workbook:Workbook,workbook_path:Path)->None:
        self._courses = {}
        self.user = user
        self.name = name
        self.workbook = workbook
        self.workbook_path = workbook_path
        self._path = None



    @property
    def ftp(self)->float:
        return self.user.ftp

    def __str__(self)->str:
        return f"CourseCollection({self.name}, {len(self._courses.keys())} Courses)"

    @property
    @pathsafe
    def path(self)->Path:
        return Path(f"output/{self.name}")


    @property
    def courses(self)->Generator[Course,None,None]:
        return self._courses.values()
    @property
    def course_names(self)->list[str]:
        return list(self._courses.keys())

    def filter(self,include:Optional[list[str]|str]=None,exclude:Optional[list[str]|str]=None)->Generator[Course,None,None]:
        filtered_names = qkfltr(list(self._courses.keys()),include=include,exclude=exclude)
        for name in filtered_names:
            yield self._courses[name]

    def build_library(self,no_save:bool=False)->None:
        courses = {d['name']:{k.title():v for (k,v) in d.items()} | {'Sport':'Bike Indoor'} \
            for d in [course.dict() for course in self]}
        
        if courses:
            existing = {}
            sheet = self.workbook.get_sheet_by_name('Library')
            rows = list(sheet.iter_rows())
            col_key = {cell.value:col+1 for (col,cell) in enumerate(rows.pop(0))}
            name_index = col_key['Name']
            max_row = 1
            for rc,row in enumerate(rows):
                if row[name_index-1].value:
                    existing[row[name_index-1].value] = rc+2
                    max_row+=1
                else:
                    break
            for course_name,course in courses.items():
                
                course_row = existing.get(course_name)
                if not course_row:
                    max_row+=1
                    course_row = max_row
                    print(f"Adding {course}")
                for key,value in course.items():
                    sheet.cell(row=course_row,column=col_key[key]).value = value
                    #print(f"Would set {(course_row,col_key[key])} to {value} for {course_name}")
            if not no_save:
                self.workbook.save(filename=self.workbook_path)

    @classmethod
    def open_excel(cls,fp:Path|str)->CourseCollection:
        reserved_sheets = ('config','schedule','library')
        output = {}
        fp = Path(fp)
        init_key = {"User Profile":("user",cls.UserProfile)}
        if fp.exists():
            name = fp.stem
            wb = open_xlsx(fp)
            config_sheet = wb['Config']
            kwargs = {"name":name}
            for named_range in config_sheet.implicit_named_ranges().values():
                if named_range.name in init_key:
                    if isinstance(init_key[named_range.name],tuple):
                        kwarg,kwarg_type = init_key[named_range.name]
                        value = kwarg_type.parse(**named_range.dict())
                    else:
                        kwarg = init_key[named_range.name]
                    kwargs[kwarg] = value 

            inst = cls(workbook=wb,workbook_path=fp,**kwargs)

            
            for name in wb.sheetnames:
                sheet = wb[name]
                if sheet.title.lower() not in reserved_sheets:
                    for course in Course.excel(inst,sheet):
                        inst._courses[course.version_name] = course
            for course in inst.courses:
                course.link()
            return inst
        else:
            raise Exception(f"File does not exist : {fp}")

    def summary(self,stats:bool=False,
                include:Optional[list[str]|str]=None,
                exclude:Optional[list[str]|str]=None)->str:
        parts = [f"{self}"]
        for course in self.filter(include,exclude):
            parts.append(course.summary(stats).replace('\n','\n\t'))           
        return '\n\t'.join(parts)

    @classmethod
    def xlsx_to_erg(cls,src:str|Path,dst:str|Path|None=None,
                    include:Optional[list[str]|str]=None,
                    exclude:Optional[list[str]|str]=None)->None:
        cc = cls.open_excel(src)
        cc.save(dst,include,exclude)

    def save(self,dst:str|Path|None,
             include:Optional[list[str]|str]=None,
             exclude:Optional[list[str]|str]=None):
        if not dst:
            dst = self.path
        else:
            dst = pathsafe(Path(dst))
        if not dst.exists():
            dst.mkdir(parents=True)

        for course in self.filter(include=include,exclude=exclude):
            course.save(dst)

    def __getitem__(self,name:str)->Course:
        return self._courses[name]

    def __iter__(self)->Generator:
        for c in self._courses.values():
            yield c


class Course:
    class CoursePowerAverages(dict):     
        def __str__(self)->str:
            return '{'+', '.join([f"{k}'@{int(v)}W" for (k,v) in self.items() if v]) + '}'

    @dataclass
    class CourseStats:
        time:float
        average:float
        np:float
        ftpif:float
        tss:int
        power_averages:CoursePowerAverages

        def __str__(self)->str:
            return ', '.join([f"{k}={v}" for (k,v) in asdict(self).items()])



    class Header(BaseModel,Templated,**Templated.class_init_kwargs):
        _key = {"Name":"name","Category":"category","Repeat":"versions","Comments":"comments"}
        _template_name = "Header"
        _key_word_argument = "header"
        _singleton = True
        name:str
        category:Category
        comments:str|None = None
        versions:str|None = None

    class PrependedCourse(BaseModel,Templated,**Templated.class_init_kwargs):
        _key = {"Name":"name","Blend Seconds":"blend"}
        _key_word_argument = "prepend"
        _template_name = "Insert Before"
        _singleton = False
        name:str
        blend:int|None = None
    class AppendedCourse(PrependedCourse,**Templated.class_init_kwargs):
        _key_word_argument = "append"
        _template_name = "Insert After"

    
    class CourseSegment(BaseModel,Templated,**Templated.class_init_kwargs):
        _key = {"Time":"time",
                "Power":"power_start",
                "Ramp-to Power":"ramp_to",
                "Exclude from last repeat":"exclude"}
        _key_word_argument = "course_data"
        _template_name = "Course"
        _singleton = False
        time:float
        power_start:float
        ramp_to:float|None = None
        exclude:bool|None = None
        start_time:float = 0.0

        @property
        def end_time(self)->float:
            return self.start_time+self.time


        @property
        def power_end(self)->int:
            if self.ramp_to:
                return self.ramp_to
            return self.power_start

        @staticmethod
        def _erg_fmt(i:float|int)->int|float:
            if i % 1 == 0:
                return int(i)
            else:
                return round(i,1)

        @property
        def power(self):
            if self.power_end:
                return mean([self.power_start,self.power_end])
            return self.power_start


        @property
        def erg(self)->str:
            return f"{self._erg_fmt(self.start_time)}\t{int(self.power_start)}\n"+\
                f"{self._erg_fmt(self.start_time+self.time)}\t{int(self.power_end)}"
        
        def __str__(self)->str:
            parts = [self._time_str(),"@",str(int(self.power_start)),'W']
            if self.ramp_to:
                parts.extend(['->',str(int(self.ramp_to)),'W'])

            return ''.join(parts)

        def _time_str(self)->str:
            minutes,min_frac = divmod(self.time,1)
            mntmp = []
            if minutes:
                mntmp.extend([str(int(minutes)),"'"])
            if min_frac:
                mntmp.extend([str(int(min_frac*60)),'"'])
      
            return ''.join(mntmp)

    _sections = {section_type._template_name:section_type for section_type in \
                (Header,PrependedCourse,AppendedCourse,CourseSegment)}

    def __init__(self,
                 collection:CourseCollection,
                 header:Header,
                 name:str,
                 version:int,
                 versioned:bool,
                 course_data:list[CourseSegment]|None = None,
                 prepend:list[PrependedCourse]|None=None,
                 append:list[AppendedCourse]|None=None)->None:
        self.collection = collection
        self.name = name
        self.category = header.category
        self.comments = header.comments
        self.version = version
        self.versioned = versioned
        self.prepend:list[Course] = []
        self.append:list[Course] = []
        self.segments:list[CourseSegment] = []
        self._prepend_names = prepend
        self._append_names = append
        self._norm_power:float|None = None
        self._stats:CourseStats|None = None

        self.linked = False
        if course_data:
            for repeat in range(self.version):               
                self.segments.extend(deepcopy(course_data))
            #recalc to make segments sortable and hashable by total time
            self._recalc_segment_time()
            for seg in list(reversed(self.segments)):
                if seg.exclude:
                    self.segments.remove(seg)
                else:
                    break
            self._recalc_segment_time()
 

    def add_segment(self,seg:CourseSegment)->None:
        if self.segments:
            seg.start_time = self.segments[-1].start_time + seg.time
        self.segments.append(seg)

    def dict(self)->dict:
        stats = asdict(self.stats)
        p_ave = {f"{k}W avg":v for (k,v) in stats.pop('power_averages').items()}
        return {'name':self.version_name,
                'category':self.category.value} | stats | p_ave |\
                    {'comments':self.comments}

    def summary(self,stats:bool=True,segments:bool=True)->str:
        parts = []
        if stats:
            parts.append(f"{self.category.value} : {self.version_name}, {self.stats}, Comments : {self.comments}")
        else:
            parts.append(f"{self.category.value} : {self.version_name}-{round(self.total_time())}', Comments : {self.comments}")
        if segments:
            for segment in self:
                parts.append(f'\t{segment}')
        return '\n'.join(parts)

    def link(self):
        if not self.linked:
            for course_link in self._prepend_names:
                self.add_linked_course(course_link,pre=True)
            for course_link in self._append_names:
                self.add_linked_course(course_link,pre=False)
            if self._append_names or self._prepend_names:
                self._recalc_segment_time()
            self.linked=True


    def add_linked_course(self,linked_course:str,pre:bool=False)->None:
        try:
            lc = self.collection[linked_course.name]
        except KeyError as ke:
            raise Exception(f"Could not find course '{linked_course.name}' to link to course '{self.name}' in "+\
                f"{','.join(list(self.collection.course_names))}.")
        if not lc.linked:
            lc.link()
        if pre:
            segs = [deepcopy(lc.segments),self.segments]
        else:
            segs = [self.segments,deepcopy(lc.segments)]

        if linked_course.blend:
            p1=segs[0][-1].power_end
            p2=segs[1][0].power_start
            if p2!=p1:
                segs[0].append(self.CourseSegment(time=round(linked_course.blend/60.00,2),
                                                  power_start=p1,
                                                  ramp_to=p2))
        self.segments = segs[0] + segs[1]

    @property
    def file_name(self)->str:
        return f"{self.name}.erg"

    @property
    def version_name(self)->str:       
        if self.version == 1 and not self.versioned:
            return self.name
        else:
            return self.name+'-'+str(self.version)+'x'

    @property
    def description(self)->str:
        return f"{self.category.value} {' '.join([k+'='+str(v) for (k,v) in self.stats.__dict__.items()])}"

    @property
    def erg(self)->str:
        parts = [f"[COURSE HEADER]",f"VERSION = {FIT_VERSION}",
                 "UNITS = ENGLISH",f"DESCRIPTION = {self.description}",
                 f"FILE NAME = {self.file_name}",
                 "FTP = 360",
                 "MINUTES WATTS",
                 "[END COURSE HEADER]","[COURSE DATA]",'\n'.join([seg.erg for seg in self]),
                 "[END COURSE DATA]"]
        return '\n'.join(parts)

    def _recalc_segment_time(self):

        for i in range(1,len(self.segments)):
            self.segments[i].start_time = self.segments[i-1].start_time + self.segments[i-1].time

    def power_by_second(self)->list[int]:
        power_by_seconds = []
        for count,seg in enumerate(self.segments):
            second_range = int(round(60*seg.time,0))
            diff = seg.power_end-seg.power_start
            for second in range(second_range):
                power_by_seconds.append(seg.power_start+second/second_range*diff)
        return power_by_seconds

    def __str__(self)->str:
        return f"Course({self.version_name}, {self.stats}, comments = {self.comments})"
    def __repr__(self)->str:
        return f"Course({self.version_name}-{round(self.total_time())}')"
    
    def total_time(self)->float:
        if self.segments:
            return round(self.segments[-1].end_time,2)
        return 0.00

    @staticmethod
    def power_for_time(t:int,pbs:list[int])->int|None:
      
        ts = 60*t
        lpbs = len(pbs)
    
        if lpbs>=ts:
            max = 0
            for c in range(lpbs - ts):
                avg = mean(pbs[c:c+ts])
                if avg > max:
                    max = avg
            return int(max)


    @property
    def stats(self)->CourseStats:
        if not self._stats:
            window_size = 30 
            total_time = self.total_time()
            pbs = self.power_by_second()
            moving_averages = []
            i = 0
            while i < len(pbs) - window_size + 1:
                this_window = pbs[i : i + window_size]
                window_average = sum(this_window) / window_size
                moving_averages.append(window_average ** 4)
                i += 1
            norm_power = round(mean(moving_averages) ** (1/4),2)
            intensity = round(norm_power / self.collection.ftp,2)
            tss = int((total_time * 60 * norm_power
                    * intensity) / (self.collection.ftp * 3600.0) * 100)

            pa = self.CoursePowerAverages()
            for average_window in (1,5,20,60):               
                pa[average_window] = self.power_for_time(average_window,pbs)

            self._stats = self.CourseStats(time = int(total_time),
                              average = round(mean(pbs),2),
                              np = norm_power,
                              ftpif = intensity,
                              tss = tss,
                              power_averages = pa)
        return self._stats


    def __iter__(self):
        for s in self.segments:
            yield s

    @property
    def path(self)->Path:
        return (Path(pathsafe(self.category)) / pathsafe(self.version_name)).with_suffix('.erg')

    def save(self,col_pth:Optional[Path]=None):
        if not col_pth:
            pth = self.collection.path / self.path
        else:
            pth = col_pth / self.path
        if not pth.parent.exists():
            pth.parent.mkdir(parents=True)
        with open(pth,'w') as f:
            print('Saving',pth)
            f.write(self.erg)

    @classmethod
    def excel(cls,collection,sheet)->list[Course]:
        ret = []       
        ranges = sheet.implicit_named_ranges()
        missing:list[str] = []
        collection_kwargs:dict[str,ls.Header|cls.PrependedCourse|cls.AppendedCourse] = {}
        course_data:dict[str,list[cls.CourseSegment]] = {}

        for req_section in (cls.Header,cls.PrependedCourse,cls.AppendedCourse,):
            try:
                for row in range(100):
                    req_section_range = ranges.pop((req_section._template_name,row),None)
                    if req_section_range:
                        break
                else:
                    raise KeyError(req_section._template_name)
            except KeyError as ke:
                missing.append(req_section._template_name)
            else:
                keys = list(req_section._key.keys())
                section_data = req_section_range.list(element=dict,
                                                          element_keys=keys)
                if req_section._singleton:
                    collection_kwargs[req_section._key_word_argument] = req_section.parse(**section_data[0])
                else:
                    collection_kwargs[req_section._key_word_argument] = [req_section.parse(**line) for line in section_data]
        if missing:
            raise Exception(f"Couldn't find data for {', '.join(missing)} in top row on sheet '{sheet.title}'.")

        header = collection_kwargs['header']

        course_ranges = list(ranges.values())
        for course_range in course_ranges:
            if course_range.name.startswith('Course'):
                suffix = ''.join([c for c in course_range.name.split('Course') if c][1:])
                if not suffix and len(course_ranges)>1:
                    suffix = ' V1'
                
                name = header.name + suffix
                r = 1
                while name in course_data:
                    if len(name_array := name.split(' '))>1:                    
                        name = f"{' '.join(name_array[:-1])} V{r}"
                    else:
                        name = f"{' '.join(name_array)} V{r}"
                    r+=1

                course_data[name] = [cls.CourseSegment.parse(**line) for line in \
                    course_range.list(element=dict,
                                      element_keys=list(cls.CourseSegment._key.keys()))]

        for course_name,course_segments in course_data.items():
            versions = cls._parse_versions(header.versions)
            for version in versions:         
                ret.append(cls(collection,
                               name=course_name,
                               version=version,
                               course_data=course_segments,
                               versioned=len(versions)>1,
                               **collection_kwargs))

        return ret

    @classmethod
    def _parse_versions(cls,ver:str|float|None)->list[str]:
        if isinstance(ver,str):
            if ver.isnumeric():
                return [int(float(ver)),]
            elif len(from_to := ver.split('-'))==2:
                start,stop = [int(float(i.strip())) for i in from_to]
                return list(range(start,stop+1))
            elif ver.find(','):
                return [int(i) for i in ver.split(',')]
            else:
                raise ValueError(ver)
        elif isinstance(ver,float):
            return [int(ver)]
        elif ver is None:
            return [1,]

