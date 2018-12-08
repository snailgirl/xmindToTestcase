import argparse
import logging
from src.main import *
from src import __version__

def main():
    parser = argparse.ArgumentParser(
        description="Convert Xmind to Excel testcases for HttpRunner.")
    parser.add_argument("-V", "--version", dest='version', action='store_true',
                        help="show version")
    parser.add_argument('xmind_file', nargs='?',
                        help="Specify Xmind file.")

    parser.add_argument('output_file', nargs='?',
                        help="Optional. Specify converted Excel testcase file.")
    args = parser.parse_args()

    if args.version:
        print("{}".format(__version__))
        exit(0)

    logging.basicConfig(level='INFO')

    xmind_file = args.xmind_file
    output_file = args.output_file

    if not xmind_file or not xmind_file.endswith(".xmind"):
        logging.error("xmind_file file not specified.")
        sys.exit(1)
    output_file_type = "xls"
    if not output_file:
        xmind_file_name = os.path.splitext(output_file)[0]
        output_file = "{}.{}.{}".format(xmind_file_name, "output", output_file_type.lower())
    else:
        output_file_suffix = os.path.splitext(output_file)[1]
        if output_file_suffix not in [".xls"]:
            logging.error("Converted file could only be in xls format.")
            sys.exit(1)
    get_xmind_content(xmind_file, output_file)

    return 0
if __name__ == '__main__':
    main()

