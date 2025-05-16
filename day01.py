# Given a signed 32-bit integer x, return x with its digits reversed. 
# If reversing x causes the value to go outside the signed 32-bit integer range [-231, 231 - 1], 
# then return 0.

# Assume the environment does not allow you to store 64-bit integers (signed or unsigned).

def reverse(self, x: int) -> int:
    t = abs(x)
    y = 0
    
    while t > 0:
        num = t % 10
        y = y * 10 + num
        t = t // 10
    
    if x < 0:
        y = -y

    if y < -2**31 or y > 2**31 - 1:
        return 0
    
    return y

        