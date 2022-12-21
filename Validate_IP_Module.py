from tkinter import messagebox

def Validate_IP(IP_Str):
    sep = IP_Str.split('.')
    if len(sep) != 4:
        # print('IP地址格式不正确')
        messagebox.showwarning('输入错误','IP地址格式不正确!')
        return False
    for i in range(4):
        try:
            # 每个元素必须为数字
            sep[i] = int(sep[i])
        except:
            # print('IP地址格式不正确')
            messagebox.showwarning('输入错误','IP地址格式不正确!')
            return False
    for i,x in enumerate(sep):
        try:
            int_x = int(x)
            if int_x < 0 or int_x > 255:
                # print('IP地址格式不正确')
                messagebox.showwarning('输入错误','IP地址格式不正确!')
                return False
            elif i == 0 and int_x == 0 :
                # print('IP地址格式不正确')
                  messagebox.showwarning('输入错误','IP地址格式不正确!')
                  return False
            elif i == 3 and int_x == 0 :
                  messagebox.showwarning('输入错误','设备IP地址不能为网络号!')
                  return False
        except ValueError as e:
            return False
    return True
