# -*- coding: utf-8 -*-
"""
PyInstaller hook for tkinterdnd2

tkinterdnd2パッケージに含まれるTkDNDライブラリを正しくバンドルします。
"""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# tkinterdnd2の全データファイル（TkDNDライブラリを含む）を収集
datas = collect_data_files('tkinterdnd2')

# サブモジュールを収集
hiddenimports = collect_submodules('tkinterdnd2')
