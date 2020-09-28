# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['auto_grant_rec_gui.py'],
             pathex=['/Users/claire/Dropbox/PG6-1/202009-grf/dist'],
             binaries=[],
             datas=[],
             hiddenimports=[],
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
          name='auto_grant_rec_gui',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
app = BUNDLE(exe,
             name='auto_grant_rec_gui.app',
             icon=None,
             bundle_identifier=None)
