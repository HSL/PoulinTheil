# PoulinTheil
An Excel add-in for prediction of partition coefficient.

Poulin, P., & Theil, F. P. (2000). A priori prediction of tissue: Plasma partition coefficients of drugs to facilitate the use of physiologically‐based pharmacokinetic models in drug discovery. _Journal of pharmaceutical sciences_, _89_(1), 16-35.

## Prerequisites
* Microsoft Windows 7 or later
* Microsoft Office 2010 or later
* Microsoft .NET 4+

## Download
Download the latest build from [releases](http://www.github.com/HSL/PoulinTheil/releases/).

Two XLL files are present in the zip archive for a release. The build (32-bit or 64-bit) of Microsoft Office installed will detemine which of the XLL files should be installed.

## Installation
To install the XLL, refer to _Add or remove an Excel add-in_ on this [Microsoft Office support page](https://support.office.com/en-us/article/Add-or-remove-add-ins-0af570c4-5cf3-4fa9-9b88-403625a0b460).

## Getting Started
Refer to Poulin & Theil to determine whether the mechanistic equations implemented in this add-in are applicable to your problem domain.

### P<sub>t:p</sub> Assuming Homogenous Distribution and Passive Diffusion

In a worksheet cell, use the Ptp_Eq11 function. The four arguments required are:

1. Species (rat, rabbit, or mouse)
2. Tissue (brain, heart, lung, muscle, skin, intestine, spleen, or bone)
3. Log vegetable oil:water partition coefficient, log K<sub>vo:w</sub>
4. Unbound fraction in lipid non-depleted plasma, fu<sub>p</sub>

Example:

```
=Ptp_Eq11("rat", "heart", 0.78, 0.92)
```

### P<sub>t:p</sub> Assuming Non-Homogenous Distribution

In a worksheet cell, use the Ptp_Eq14 function. The two arguments required are:

1. Tissue (brain, heart, lung, muscle, skin, intestine, spleen, or bone)
2. Unbound fraction in lipid non-depleted plasma, fu<sub>p</sub>

Example:

```
=Ptp_Eq14("skin", 0.22)
```