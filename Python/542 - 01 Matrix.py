class Solution:
    def updateMatrix(self, mat: List[List[int]]) -> List[List[int]]:
        m = len(mat)
        n = len(mat[0])
        
        seen = set()
        q = deque()
        
        for i in range(m):
            for j in range(n):
                if mat[i][j] == 0:
                    seen.add((i,j))
                    q.append([i,j])
        
        while q:
            for i in range(len(q)):
                x, y = q.popleft()
                for (i , j) in ((x-1, y), (x, y-1), (x+1, y), (x, y+1)) :
                    if 0 <= i < m and 0 <= j < n and (i, j) not in seen :
                            mat[i][j] = mat[x][y] + 1
                            seen.add((i,j))
                            q.append([i,j])
        return mat
       
#O(m*n) time and space 
