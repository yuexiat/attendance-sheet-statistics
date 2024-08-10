def column_letter_to_number(column_letter):
    column_number = 0
    factor = 1

    for i in range(len(column_letter) - 1, -1, -1):
        char = column_letter[i]
        column_number += (ord(char.upper()) - ord('A') + 1) * factor
        factor *= 26



    return column_number

print(column_letter_to_number('n'))