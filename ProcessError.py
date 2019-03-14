class ProcessError(Exception):
  def __init__(self, info,filter):
      self.info = info
      self.filter=filter