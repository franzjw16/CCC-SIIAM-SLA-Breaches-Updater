import sys
sys.setrecursionlimit(5000)

# -*- mode: python -*-

block_cipher = None


a = Analysis(['sla_breaches.py'],
             pathex=['D:\\AVC_Washing\\proactive\\vm version\\static\\webform_csv'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=['jedi', 'PyQt5', 'PIL', 'sqlalchemy', 'pandas', 'numpy', 'matplotlib'],
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
          name='CCC SIIAM SLA Breaches Spreadsheet Updater',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
