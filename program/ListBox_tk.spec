# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['ListBox_tk.py', 'SeatNumber.py', 'Sort_Students.py', 'settings.py'],
             pathex=['D:\\Eisu_Seat_manage\\program'],
             binaries=[],
             datas=[],
             hiddenimports=['openpyxl','pkg_resources.py2_warn'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='ListBox_tk',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
