from func import update_sheet, process_send

linha = int(input("Qual a linha desejada: "))
coluna = int(input("Qual a coluna desejada: "))

update_sheet("Reservar 2.xlsx",linha=linha, coluna=coluna)
process_send()