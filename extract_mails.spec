# -*- mode: python -*-

block_cipher = None


a = Analysis(['extract_mails.py'],
             pathex=['Z:\\Projects\\MailExtractor'],
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
		Tree('antiword', prefix='antiword\\'),
          a.zipfiles,
          a.datas,
          [],
          name='extract_mails',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=True )
