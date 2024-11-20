# Part 1

class PasswordValidator:
    def __init__(self, selected_criteria):
        self.selected_criteria = selected_criteria
        self.allowed_special_chars = "!@#"

    def validate(self, password):
        if len(password) < 8:
            return f"'{password}' is invalid. Reason: Less than 8 characters."

        invalid_reasons = []

        if 1 in self.selected_criteria and not self.has_uppercase(password):
            invalid_reasons.append("Missing Uppercase letters")
        if 2 in self.selected_criteria and not self.has_lowercase(password):
            invalid_reasons.append("Missing Lowercase letters")
        if 3 in self.selected_criteria and not self.has_numbers(password):
            invalid_reasons.append("Missing Numbers")
        if 4 in self.selected_criteria and not self.has_special_chars(password):
            invalid_reasons.append("Missing Special characters")

        invalid_chars = self.get_invalid_special_chars(password)
        if invalid_chars:
            invalid_reasons.append(f"Contains invalid special characters: {', '.join(invalid_chars)}")

        if invalid_reasons:
            return f"'{password}' is invalid. Reason: {', '.join(invalid_reasons)}"
        return f"'{password}' is a valid password."

    def has_uppercase(self, password):
        return any(char.isupper() for char in password)

    def has_lowercase(self, password):
        return any(char.islower() for char in password)

    def has_numbers(self, password):
        return any(char.isdigit() for char in password)

    def has_special_chars(self, password):
        return any(char in self.allowed_special_chars for char in password)

    def get_invalid_special_chars(self, password):
        return [char for char in password if not char.isalnum() and char not in self.allowed_special_chars]

"""# Main Function"""

# Main function to drive the program
def main():
    print("Select the criteria to check passwords:")
    print("1. Uppercase letters (A-Z)")
    print("2. Lowercase letters (a-z)")
    print("3. Numbers (0-9)")
    print("4. Special characters (!, @, #)")
    user_criteria = input("Enter the numbers of the criteria you want to check (e.g., 1,2,3): \n\n")
    selected_criteria = {int(x) for x in user_criteria.split(',')}

    # password_list = [
    #     "abc12345",
    #     "abc",
    #     "123456789",
    #     "abcdefg$",
    #     "abcdefgABHD!@313",
    #     "abcdefgABHD$$!@313"
    # ]

    validator = PasswordValidator(selected_criteria)

    password = input("Enter the password to validate: \n\n")
    result = validator.validate(password)
    print(result)
    # for password in password_list:
    #     result = validator.validate(password)
    #     print(result)

# Execute the main function
main()

#Part 2

with open("/content/Input.txt", "r") as file:
    lines = file.readlines()
    s = ""
    for line in lines:
        s += line
    lines = s.split("\n")

# Main function to drive the program
def main():
    print("Select the criteria to check passwords:")
    print("1. Uppercase letters (A-Z)")
    print("2. Lowercase letters (a-z)")
    print("3. Numbers (0-9)")
    print("4. Special characters (!, @, #)")
    user_criteria = input("Enter the numbers of the criteria you want to check (e.g., 1,2,3): \n\n")
    selected_criteria = {int(x) for x in user_criteria.split(',')}

    # password_list = [
    #     "abc12345",
    #     "abc",
    #     "123456789",
    #     "abcdefg$",
    #     "abcdefgABHD!@313",
    #     "abcdefgABHD$$!@313"
    # ]

    validator = PasswordValidator(selected_criteria)

    for password in lines:
        result = validator.validate(password)
        print(result)

# Execute the main function
main()