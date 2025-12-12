from backend.add_child_handler import ParentChildHandler

# Initialize handler
handler = ParentChildHandler()

# Step 1: Send OTP to child email
print("\n--- Sending OTP ---")
send_response = handler.send_child_otp({
    "parent_email": "parent@example.com",
    "child_email": "child@example.com"
})
print(send_response)

# Step 2: Once you get OTP in child's email, enter it here to verify
otp_code = input("\nEnter the OTP received in child's email: ")

# Step 3: Verify OTP and add child
print("\n--- Verifying OTP ---")
verify_response = handler.verify_child_otp({
    "parent_email": "parent@example.com",
    "child_email": "child@example.com",
    "otp": otp_code
})
print(verify_response)
