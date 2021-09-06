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


If ```RDCOMClient``` fails or if Echoview is crashing more frequently then exprected when using EchoviewR:

Please install RDCOMClient through:

```{r altRDCOMClient, eval=FALSE}
devtools::install_github("dkyleward/RDCOMClient")
```

In older versions of ```R``` it was recommended to install RDCOMClient through:  
```{r dependPacks, eval=FALSE}
devtools::install_github("omegahat/RDCOMClient")
```

If the install still fails because of ```Rtools``` go to [https://cran.r-project.org/bin/windows/Rtools/] and install the latest version of ```Rtools```. Once Rtools is installed, you have to restart RStudio and then install ```RDCOMClient``` before proceeding to install ```EchoviewR```.

You are then ready to install the ```EchoviewR``` package directly from github using ```devtools``` (Wickham & Chang, 2016):


```{r install,eval=FALSE}
if (!requireNamespace("devtools",quietly=TRUE)) install.packages("devtools")
# Latest devtools syntax
devtools::install_github("AustralianAntarcticDivision/EchoviewR", build_opts = c("--no-resave-data", "--no-manual"))
## For devtools v < 2.0
# devtools::install_github("AustralianAntarcticDivision/EchoviewR", build_vignettes = TRUE, force_deps=TRUE)
```

NB if you would like the older ```EchoviewR``` version (v1.0) that was used during the 2017 meeting of the CCAMLR Subgroup on Acoustics, Survey and Analysis Methods, please use the following code:

```{r installASAM,eval=FALSE}
devtools::install_github('AustralianAntarcticDivision/EchoviewR',ref='v1.0')
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
