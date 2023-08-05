from __future__ import annotations
from typing import Optional,Any,Generator,Callable,Iterator,Iterable

def qkfltr(s:Iterable[str],
            include:Optional[str|list[str]]=None,
            exclude:Optional[str|list[str]]=None,
            case_sens:bool=False,
            limit:int=0)->list:
    def _val(i:Iterable)->list[str]:
        if isinstance(i,str):
            return [i,]
        elif isinstance(i,Iterable):
            btypes = [type(e) for e in i if not isinstance(e,str)]
            if not btypes:
                return list(i)
            raise TypeError(f"Every element of 's' must be of type <str> not {','.join(btypes)}")
        raise TypeError(f"'search' must be an iterable, not {type(i)}")

    if not case_sens:
        _case = lambda x:x.lower()
        for iterable_arg in (include,exclude):
            if iterable_arg:
                iterable_arg = [e.lower() for e in iterable_arg]
    else:
        _case = lambda x:x
    
    included = []
    excluded = set([])

    if not limit:
        limit = len(s)

    if include:
        include = _val(include)
        include_lmds:list[tuple[Callable,str]] = []

        for inc_element in include:
            if inc_element.startswith('*') and inc_element.endswith('*'):
                include_lmds.append((lambda a,b:_case(a).find(b[1:-1])!=-1,inc_element))
            elif inc_element.startswith('*'):
                include_lmds.append((lambda a,b:_case(a).endswith(b[1:]),inc_element))
            elif inc_element.endswith('*'):
                include_lmds.append((lambda a,b:_case(a).startswith(b[:-1]),inc_element))
            else:
                include_lmds.append((lambda a,b:_case(a)==b,inc_element))
        
        res_len = 0
        for element in s:
            for inc_lmb,t in include_lmds:
                if inc_lmb(element,t):
                    res_len+=1
                    included.append(element)
                    break
            if res_len>=limit:
                break
    else:
        included = list(s)

    if exclude:
        exclude = _val(exclude)
        exclude_lmds:list[tuple[Callable,str]] = []
        
        for ex_element in exclude:
            if ex_element.startswith('*') and ex_element.endswith('*'):
                exclude_lmds.append((lambda a,b:_case(a).find(b[1:-1])!=-1,ex_element))
            elif ex_element.startswith('*'):
                exclude_lmds.append((lambda a,b:_case(a).endswith(b[1:]),ex_element))
            elif ex_element.endswith('*'):
                print(ex_element[:-1])
                exclude_lmds.append((lambda a,b:_case(a).startswith(b[:-1]),ex_element))
            else:
                exclude_lmds.append((lambda a,b:_case(a)==b,ex_element))
        for element in included:
            for ex_lmb,t in exclude_lmds:
                if ex_lmb(element,t):
                    excluded.add(element)
                    break
        included = [e for e in included if e not in excluded]  
 
    return included