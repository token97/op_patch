# op_patch
**Python 2.7**

**Openpyxl 2.3.3** patch  
Allow to load images when reading existing excel file



# usage :
```python
from openpyxl import load_workbook

from op_patch.excel import load_workbook as patch_load_workbook
load_workbook = patch_load_workbook
```

Thanks to [openpyxl](https://openpyxl.readthedocs.org/en/2.3.3/) team and thanks to [Sergey Pikhovkin](https://gist.github.com/pikhovkin) for her first version of the [patch](https://gist.github.com/pikhovkin/543709a2e2827d9c345d).