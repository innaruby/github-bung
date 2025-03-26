    self.value = value
    ^^^^^^^^^^
  File \openpyxl\cell\cell.py", line 218, in value
    self._bind_value(value)
  File "\openpyxl\cell\cell.py", line 187, in _bind_value
    raise ValueError("Cannot convert {0!r} to Excel".format(value))
ValueError: Cannot convert <NA> to Excel
