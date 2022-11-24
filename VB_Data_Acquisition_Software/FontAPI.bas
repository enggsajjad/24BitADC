Attribute VB_Name = "FontAPI"
'API declarations font handling
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Public Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type

Public Enum FontWeights
  FW_DONTCARE = 0
  FW_THIN = 100
  FW_EXTRALIGHT = 200
  FW_LIGHT = 300
  FW_NORMAL = 400
  FW_MEDIUM = 500
  FW_SEMIBOLD = 600
  FW_BOLD = 700
  FW_EXTRABOLD = 800
  FW_HEAVY = 900
End Enum

Public Enum FontCharSets
  ANSI_CHARSET = 0
  DEFAULT_CHARSET = 1
  SYMBOL_CHARSET = 2
  MAC_CHARSET = 77
  SHIFTJIS_CHARSET = 128
  HANGEUL_CHARSET = 129
  CHINESEGIG5_CHARSET = 136
  OEM_CHARSET = 255
End Enum

Public Enum FontOutPrecisions
  OUT_DEFAULT_PRECIS = 0
  OUT_STRING_PRECIS = 1
  OUT_CHARACTER_PRECIS = 2
  OUT_STROKE_PRECIS = 3
  OUT_TT_PRECIS = 4
  OUT_DEVICE_PRECIS = 5
  OUT_RASTER_PRECIS = 6
  OUT_TT_ONLY_PRECIS = 7
  OUT_OUTLINE_PRECIS = 8
End Enum

Public Enum FontClipPrecisions
  CLIP_DEFAULT_PRECIS = 0
  CLIP_CHARACTER_PRECIS = 1
  CLIP_STROKE_PRECIS = 2
  CLIP_LH_ANGLES = 16
  CLIP_TT_ALWAYS = 32
  CLIP_EMBEDDED = 128
  CLIP_TO_PATH = 4097
End Enum

Public Enum FontQuality
  DEFAULT_QUALITY = 0
  DRAFT_QUALITY = 1
  PROOF_QUALITY = 2
End Enum

Public Enum FontPitch
  DEFAULT_PITCH = 0
  FIXED_PITCH = 1
  VARIABLE_PITCH = 2
End Enum

Public Enum FontFamily
  FF_DONTCARE = 0
  FF_ROMAN = 16
  FF_SWISS = 32
  FF_MODERN = 48
  FF_SCRIPT = 64
  FF_DECORATIVE = 80
End Enum

