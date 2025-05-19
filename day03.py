# The Fibonacci numbers, commonly denoted F(n) form a sequence, 
# called the Fibonacci sequence, such that each number is the sum of the two preceding ones, starting from 0 and 1. That is,

class Solution:
    def fib(self, n: int) -> int:
        a = 0
        b = 1
        c = 1
        if n == 0:
            return 0
        elif n == 1:
            return 1
        else: 
            while n > 1:
                c = a + b
                a = b
                b = c
                n = n - 1
            return c