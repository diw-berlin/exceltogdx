# Excel to gdx
A simple python tool to extract sets and parameters from excel files by creating GDX files for GAMS

See sheet name 'py' in the excel files that describes sets and parameters settings. Similar approach for https://www.gams.com/latest/docs/T_GDXXRW.html

Particularly useful for linux and ios where GDXXRW.exe is not compatible, see here: https://forum.gamsworld.org/viewtopic.php?t=10418#p24019

### NEW FEATURE (v0.0.5):+INF, -INF, EPS

### NEW FEATURE (v0.0.6): gams_dir as new argument. 

Required library:
 - gdxpds
 - pandas
 - numpy
 - openpyxl
 - GAMS API for python
 
Installation
 
    pip install exceltogdx
 
Delete
 
     pip uninstall exceltogdx

See notebook [example.ipynb](https://github.com/diw-berlin/exceltogdx/tree/master/notebooks) and download the excel file [test.xlsx](https://github.com/diw-berlin/exceltogdx/tree/master/notebooks)

```
from exceltogdx import exceltogdx
frames = exceltogdx('test.xlsx','test.gdx', gams_dir=None)
```

Repository: https://github.com/diw-berlin/exceltogdx


If you find and error in the code, please raise an issue or solve it by pushing a commit.
