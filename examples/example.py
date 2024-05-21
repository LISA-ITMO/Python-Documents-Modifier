import sys, os

parent_dir = os.path.abspath(os.path.join(os.path.dirname('.'), "."))
sys.path.append(parent_dir)

from src.odt.odtredactor import ODTRedactor

if __name__ == '__main__':

    file = ODTRedactor(
        path_input=r'examples/example_file.odt',
        path_output=r'examples/example_file_edited.odt'
    )
    
    file.add_comment_by_text("вышедший", "test", author="Lisa")

    file.save_file()