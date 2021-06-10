excels2vensim
=============

Systems Dynamics models often use large amounts of input data. Reading this data is complicated when it corresponds to multidimensional matrices. In the case of Vensim, if the data is read from '.xlsx' files, the data can only be read in two-dimensional tables at most, which implies the introduction of multiple equations in the model file.

The excels2vensim library aims to simplify the incorporation of equations from external data into Vensim. Given the name of the variable and information on how its dimensions are distributed, the library returns the equations for copying and pasting into the Vensim model. In addition, the library uses cellrange names to write the equations, and automatically creates the cellranges in the '.xlsx' file.

The automation of this process can save a lot of time for the modeller, as well as reduce possible human error by avoiding entering a large number of equations by hand. The library is also flexible and works for any number of dimensions. In addition, it allows reading data of which one dimension is spread over different files or sheets.

This library is able to automate the generation of the following type of equations:
- GET XLS CONSTANTS
- GET DIRECT CONSTANTS
- GET XLS DATA (with or without keywords)
- GET DIRECT DATA (with or without keywords)
- GET XLS LOOKUPS
- GET DIRECT LOOKUPS

## Requirements
- Python 3.7+
- PySD 1.4+

## Resources
See the [project documentation](http://excels2vensim.readthedocs.org/) for information about:

- [Installation](http://excels2vensim.readthedocs.org/en/latest/installation.html)
- [Usage](http://excels2vensim.readthedocs.org/en/latest/usage.html)

Some jupyter notebook examples are available in the [examples folder of the GitLab repository](https://gitlab.com/eneko.martin.martinez/excels2vensim/-/tree/master/examples).

## Authority and acknowledgmentes
This library was originally developed for [H2020 LOCOMOTION](https://www.locomotion-h2020.eu/) project by [@eneko.martin.martinez](https://gitlab.com/eneko.martin.martinez) at [Centre de Recerca Ecològica i Aplicacions Forestals (CREAF)](http://www.creaf.cat/).

This project has received funding from the European Union’s Horizon 2020
research and innovation programme under grant agreement No. 821105.
