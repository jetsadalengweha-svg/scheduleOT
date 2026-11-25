class Solution:
    def solveSudoku(self, board: List[List[str]]) -> None:
        self.solve(board)

    def solve(self, board):
        for i in range(9):
            for j in range(9):
                if board[i][j] == '.':
                    for ch in '123456789':
                        if self.is_valid(board, i, j, ch):
                            board[i][j] = ch
                            if self.solve(board):
                                return True
                            board[i][j] = '.'
                    return False
        return True

    def is_valid(self, board, row, col, ch):
        for i in range(9):
            if board[row][i] == ch or board[i][col] == ch:
                return False
            r = 3 * (row // 3) + i // 3
            c = 3 * (col // 3) + i % 3
            if board[r][c] == ch:
                return False
        return True
