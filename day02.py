# Given an integer x, return true if x is a palindrome, and false otherwise.

class Solution:
    def isPalindrome(self, x: int) -> bool:
        i = 0
        j = len(str(x))-1
        s = str(x)

        while i < j:
            if s[i] != s[j]:
                return False
            else:
                i+=1
                j-=1
        
        return True

        

# Given an array of integers nums and an integer target, return indices of the two numbers such that they add up to target.

class Solution:     
    def twoSum(self, nums: List[int], target: int) -> List[int]:
        one = 0
        two = 0

        for i in range(len(nums)):
            t = target - nums[i]
            if t in nums and nums.index(t) != i:
                return [i, nums.index(t)]
                
        return [-1, -1]