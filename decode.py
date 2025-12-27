# виженер
alphabet = ['А', 'Б', 'В', 'Г', 'Д', 'Е', 'Ё', 'Ж', 'З', 'И', 'Й', 'К', 'Л', 'М', 'Н',
            'О', 'П', 'Р', 'С', 'Т', 'У', 'Ф', 'Х', 'Ц', 'Ч', 'Ш', 'Щ', 'Ъ', 'Ы', 'Ь',
            'Э', 'Ю', 'Я', 'а', 'б', 'в', 'г', 'д', 'е', 'ё', 'ж', 'з', 'и', 'й', 'к',
            'л', 'м', 'н', 'о', 'п', 'р', 'с', 'т', 'у', 'ф', 'х', 'ц', 'ч', 'ш', 'щ',
            'ъ', 'ы', 'ь', 'э', 'ю', 'я', '0', '1', '2', '3', '4', '5', '6', '7', '8',
            '9', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
            'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'a', 'b', 'c',
            'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r',
            's', 't', 'u', 'v', 'w', 'x', 'y', 'z', ' ', '.', ',', ';', ':', '!', '?',
            '(', ')', '/', '№', '@', '#', '$', '%', '*', '=', '+'
            ]


def choice(text, keyy):

    x = False
    while x is False:
        option = input("введите опцию (k/d): ")
        if option == 'k':
            to_code(text, keyy)
            x = True
        elif option == 'd':
            to_decode(text, keyy)
            x = True
        else:
            continue


def make_long_key(text, keyy):

    clear_key = ""
    for k_letter in keyy:
        if k_letter not in alphabet:
            k_letter = " "
        clear_key += k_letter

    key_str = ""
    while len(key_str) < len(text):
        key_str += clear_key
    long_key = key_str[:len(text)]

    # print(long_key)

    return long_key


def to_code(text, keyy):
    the_code = ""
    long_key = make_long_key(text, keyy)

    for i, letter in enumerate(text):
        if letter not in alphabet:
            letter = " "

        text_letter_num = alphabet.index(letter)+1
        long_key_letter_num = alphabet.index(long_key[i])
        moved_text_letter_num = text_letter_num + long_key_letter_num

        if moved_text_letter_num > len(alphabet):
            code_letter_num = moved_text_letter_num-len(alphabet)
        else:
            code_letter_num = text_letter_num+long_key_letter_num

        code_letter = alphabet[code_letter_num-1]
        the_code += code_letter

        # print(letter, long_key_letter_num, code_letter)

    print(the_code)


def to_decode(text, keyy):
    the_decode = ""
    long_key = make_long_key(text, keyy)

    for i, letter in enumerate(text):
        if letter not in alphabet:
            letter = " "

        text_letter_num = alphabet.index(letter)+1
        long_key_letter_num = alphabet.index(long_key[i])
        moved_text_letter_num = text_letter_num-long_key_letter_num

        if moved_text_letter_num < 1:
            code_letter_num = len(alphabet)+moved_text_letter_num
        else:
            code_letter_num = text_letter_num-long_key_letter_num

        code_letter = alphabet[code_letter_num-1]
        the_decode += code_letter

        # print(letter, -long_key_letter_num, code_letter)

    print(the_decode)


# text = "Если ты хочешь исключить зависимость от пула, можно самому держать валидаторский узел."
# keyy = "бутылка ванючий жопа сраный мяу==="

text = input("введите сообщение: ")
keyy = input("введите ключ: ")

choice(text, keyy)
