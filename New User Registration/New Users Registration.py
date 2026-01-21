import json

login_option={}

while True:
    try:
        user_no = int(input('Enter No. of Users to Register : '))
        break
    except:
        print("Invalid input! Please enter a number.")

for i in range(1,user_no+1):
    print('\n')
    user_name = input(f'Enter name of user {i} : ').capitalize()
    user_api_key = input(f'Enter API KEY of user {i} : ')
    user_api_secret = input(f'Enter API SECRET of user {i} : ')
    user_api_auth = input(f'Enter API AUTH-CODE of user {i} : ')

    while True:
        try:
            user_pin = int(input(f'Enter PIN of user {i}: '))
            break
        except ValueError:
            print("PIN must be numbers only.")

    while True:
        user_mob = input(f'Enter Mobile No. of user {i}: ')
        if user_mob.isdigit() and len(user_mob) == 10:
            user_mob = int(user_mob)   # convert to integer only after validation
            break
        else:
            print("Mobile number must be exactly 10 digits.")
            
    user_fullname = input(f'Enter Full Name of user {i} : ')
    login_option[user_name] = {'api_key':user_api_key, 'api_secret':user_api_secret, 'api_auth':user_api_auth, 'pin':user_pin, 'Mob No.':user_mob, 'full_name':user_fullname}

with open('login_details.json', 'w') as file_write:
    json.dump(login_option, file_write, indent=4)

print('\n')
print("User login details saved to login_details.json")
