# Validate Port Number    [ April 25, 2024 ]
# Written by Vincent.

def validate_port_num(port_num):
    try: 
        int(port_num)
        check_int_result = True
    except ValueError: 
        check_int_result = False

    if check_int_result == True:
        if int(port_num) in range(1, 65536):
            return True
        else:
            return False
    else:
        return False
