class Solution(object):
    def maxProfit(self, prices):
        """
        :type prices: List[int]
        :rtype: int
        """
        profit = 0
        price = float('inf')
        for i in range(len(prices)): 
            if prices[i] < price:
                price = prices[i]
            elif prices[i] - price > profit:
                profit = prices[i] - price
        
        return profit
        
##O(n) time 
##O(n) space
