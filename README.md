# EchoviewR
The R package EchoviewR - a free interface between Echoview (R) and R using COM scripting.

Copyright and Licence

    Copyright (C) 2015 Lisa-Marie Harrison, Martin Cox, Georg Skaret and Rob Harcourt.
    
    This file is part of EchoviewR.
    
    EchoviewR is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.
    
    EchoviewR is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.
    
    You should have received a copy of the GNU General Public License
    along with EchoviewR.  If not, see <http://www.gnu.org/licenses/>.


This package is open for community development and we encourage users to extend the package as they need. We are not liable for any losses when using EchoviewR. 

If using EchoviewR, please cite as:

```{r citation}
citation('EchoviewR')
```


### Installing EchoviewR
Currently, ```EchoviewR``` has quite a few dependencies, something we are hoping to rectify.  Meanwhile, you will need to install the following ```R``` packages:

```sp```,```geosphere```,```maptools```,```RDCOMClient``` and ```lubridate```
  
You can install all these packages with the following  ```R``` code:

```{r dependPacks,eval=FALSE}
install.packages(c('sp','geosphere','maptools','RDCOMClient','lubridate'))
```

You are then ready to install the ```EchoviewR``` package directly from github using ```devtools``` (Wickham & Chang, 2016):


```{r install,eval=FALSE}
if (!requireNamespace("devtools",quietly=TRUE)) install.packages("devtools")
 devtools::install_github('AustralianAntarcticDivision/EchoviewR')
```

You should then be ready to work with EchoviewR:
```{r startEVR, eval=FALSE}
library(EchoviewR)
```

### Getting started

We are currently updating the ```EchoviewR``` vignettes.  Once available, the vignettes will use data available at the Australian Antarctic Division Data Centre [doi: http://dx.doi.org/10.4225/15/54CF081FB955F].

### Acknowledgements

The authors would like to thank Echoview (R) for their help and support during the development of this R package.

Echoview (R) is a registered trademark of Echoview Software Pty Ltd. Website: www.echoview.com/

### References

Hadley Wickham and Winston Chang (2017). devtools: Tools to Make Developing R Packages Easier. R package version 1.34.4.
  https://CRAN.R-project.org/package=devtools
