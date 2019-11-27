# excel to gdx
A simple python tool to extract sets and parameters from excel files by creating GDX files for GAMS

See sheet name 'py' in the excel files that describes sets and parameters settings. Similar approach for https://www.gams.com/latest/docs/T_GDXXRW.html

Particularly useful for linux and ios where GDXXRW.exe is not compatible, see here: https://forum.gamsworld.org/viewtopic.php?t=10418#p24019

Required library:
 - gdxpds
 - pandas
 - openpyxl
 - GAMS API for python
 
Installation
 
    pip install exceltogdx
 
Delete
 
     pip uninstall exceltogdx

<font color='red'>If you find and error in the code, please raise an issue or solve it by pushing a commit.</font>

See notebook example.ipynb
