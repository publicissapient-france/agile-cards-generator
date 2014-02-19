import os

def main():
    file_name = prepare_output_file('output', 'xlsx')
    print(file_name)

def prepare_output_file(output_file, extension):
    file_name = None
    if output_file is not None:
        file_name = output_file
    else:
        file_name = 'output.' + extension
    output_path = 'output'
    if not os.path.isdir(output_path):
        os.makedirs(output_path, exist_ok=True)
    file_name = os.path.join(output_path, file_name)
    if os.path.isfile(file_name):
        os.remove(file_name)
    return file_name


if __name__ == "__main__":
    main()