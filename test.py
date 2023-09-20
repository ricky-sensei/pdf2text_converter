import re


def add_newline_after_period(text):
    # 句点の後に「」または『』がある場合、その後に改行を追加する
    modified_text = re.sub(r"([。])([」』]?)", r"\1\2\n", text)
    return modified_text


# テスト用の文章
input_text = "「こんにちは。」とたかしは言った。これはテストです。こんにちは。彼は言った。"

# 関数を呼び出して改行を追加した文章を取得
result_text = add_newline_after_period(input_text)

# 結果を表示
print(result_text)
