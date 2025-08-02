class response_result:
    def __init__(self, code: int, message: str = None, data: dict = None):
        self.code = code
        self.message = message
        self.data = data if data is not None else {}

    def to_dict(self):
        return {
            'code': self.code,
            'message': self.message,
            'data': self.data
        }

    def __str__(self):
        return f"ResponseResult(code={self.code}, message={self.message}, data={self.data})"