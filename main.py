import argparse
from trainercourses.course import CourseCollection
from pathlib import Path
import os

def prompt(o:str)->bool:
    i = input(f"{o}\n->")
    if i.lower().startswith('y'):
        return True
    return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('-inc','--include_courses', type=str.lower,nargs='+',required=False)
    parser.add_argument('-exc','--exclude_courses', type=str.lower,nargs='+',required=False)
    parser.add_argument('-src','--source',default='',type=Path)
    parser.add_argument('-dst','--destination',default='',type=Path)
    parser.add_argument('-p','--print',action="store_true",
                        help="Print all courses to console.")

    parser.add_argument('-e','--export',action="store_true",
                        help="Export to folder structure in parent directory or in 'src' directory if supplied")
    parser.add_argument('-f','--format',default='fit',
                        help="Export to what format?")

    args = parser.parse_args()

    cc = CourseCollection.open_excel(args.source)

    if args.print:
        print(cc.summary(stats=True))
    if args.export:
        if not args.destination:
            dst = Path(os.getcwd())
        else:
            dst = Path(os.path.abspath(args.destination))
        if not dst.exists():
            if prompt(f"{dst} does not exist. Create it?"):
                dst.mkdir(parents=True)
        cc.save(dst=dst,include=args.include_courses,exclude=args.exclude_courses)