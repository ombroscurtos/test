import pyscreenshot as ImageGrab
from win32 import win32clipboard
from io import BytesIO

# fullscreen
def tirar_print():
              #cima#dir#baixo
  regiao = (43,282,760,530)
  im = ImageGrab.grab(regiao)
  #im.show()
  
  output = BytesIO()
  im.convert('RGB').save(output, 'BMP')
  data = output.getvalue()[14:]
  output.close()

  win32clipboard.OpenClipboard()
  win32clipboard.EmptyClipboard()
  win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
  win32clipboard.CloseClipboard()






