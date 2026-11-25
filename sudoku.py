board = [
    list("53..7...."),
    list("6..195..."),
    list(".98....6."),
    list("8...6...3"),
    list("4..8.3..1"),
    list("7...2...6"),
    list(".6....28."),
    list("...419..5"),
    list("....8..79"),
]

Solution().solveSudoku(board)

for row in board:
    print(" ".join(row))

# ตรวจว่าผลลัพธ์ถูกต้อง
full = set("123456789")
ok = (
    all(set(r) == full for r in board)
    and all({board[r][c] for r in range(9)} == full for c in range(9))
    and all(
        {board[br + i][bc + j] for i in range(3) for j in range(3)} == full
        for br in (0, 3, 6) for bc in (0, 3, 6)
    )
)
print("ถูกต้อง" if ok else "ผิด")




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
