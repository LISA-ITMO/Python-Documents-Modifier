from odtredactor import ODTRedactor

if __name__ == '__main__':
    file = ODTRedactor(
        path_input='test.odt',
        path_output='test1.odt'
    )
    file.add_comment_by_text("Рукопись была скептически", 'комментарий')