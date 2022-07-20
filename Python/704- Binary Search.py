class Solution:
    def search(self, nums: List[int], target: int) -> int:
        left = 0
        right = len(nums)-1
        
        while left <= right:
            n = (right + left)//2
            if target == nums[n]:
                return n
            elif target < nums[n]:
                right = n - 1
            elif target > nums[n]:
                left = n + 1 
        return -1
#O(log n) time && O(1) space 
