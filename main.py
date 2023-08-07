import os
import sys
import argparse
from trainercourses.course import CourseCollection
from pathlib import Path


def prompt(o:str)->bool:
    i = input(f"{o}\n->")
    if i.lower().startswith('y'):
        return True
    return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--include', type=str.lower,nargs='+',required=False,help="Courses to include.  Example: *course name* *another course name*")
    parser.add_argument('--exclude', type=str.lower,nargs='+',required=False,help="Courses to exclude.  Example: *course name* *another course name*")
    parser.add_argument('--src',type=Path,help='Path (relative or absolute) to xlsx course collection.')
    parser.add_argument('--dst',type=Path,help='Directory (relative or absolute) for exports.')

    parser.add_argument('-p','--print',action="store_true",
                        help="Print all courses to console.")
    parser.add_argument('-b','--build',action="store_true",
                        help="Update library sheet in xlsx file.")
    parser.add_argument('-e','--export',action="store_true",
                        help="Export to folder structure in $/export or in 'dst' directory if supplied")
    parser.add_argument('-f','--format',default='erg',
                        help="Export to what format?")
   
    args = parser.parse_args()
    
    if not args.src:
        source = Path(os.getcwd()) / 'Collection.xlsx'
        
    else:
        source = Path(os.path.abspath(args.src))

    if not source.exists():
        raise ValueError(f"Source not found.  Either specify a custom path to a Course Collection .xlsx file or make sure "+\
            f"the default file of $\Collection.xlsx exists.")
        sys.exit()

    cc = CourseCollection.open_excel(source)
    inc,exc = "",""
    if args.include:
        inc = f" (include: {','.join(args.include)})"
    if args.exclude:
        exc = f" (exclude: {','.join(args.exclude)})"
    if args.print:
        print(f'Printing Courses...{inc}{exc}')
        print(cc.summary(stats=True,include=args.include,exclude=args.exclude))
    if args.build:
        print('Updating Library tab...')
        cc.build_library()
    if args.export:
        print(f'Exporting Courses...{inc}{exc}')
        if not args.dst:
            dst = Path(os.getcwd()) / f"export"
        else:
            dst = Path(os.path.abspath(args.dst))
        if not dst.exists():
            if prompt(f"{dst} does not exist. Create it?"):
                dst.mkdir(parents=True)
            else:
                sys.exit()
        cc.save(dst=dst,include=args.include,exclude=args.exclude)
sys.exit()