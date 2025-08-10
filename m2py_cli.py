import argparse
from m2py_core import convert_m_to_python

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', required=True, help='Input .m file')
    parser.add_argument('--output', required=True, help='Output .py file')
    args = parser.parse_args()

    with open(args.file, 'r') as f:
        m_code = f.read()
    py_code = convert_m_to_python(m_code)
    with open(args.output, 'w') as f:
        f.write(py_code)

if __name__ == '__main__':
    main()