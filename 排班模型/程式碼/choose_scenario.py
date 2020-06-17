wrong_ans = True
while(wrong_ans):
    user_choice = input("選擇需要的班表種類 (請輸入0到2的數字):\n0: 一般班表\n1: 情境3 (6天休1天)\n2: 情境4 (北中高三地輪流休一個假日)\n")
    if user_choice == '0':
        import project_main
        break
    elif user_choice == '1':
        import scen3
        break
    elif user_choice == '2':
        import scen4
        break
    else:
        print('請輸入0到2的數字\n')
print('新班表完成！')