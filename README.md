# EliteQuant_Excel
Quant Modelling, portfolio management and Trading Platform in Excel

* [Platform Introduction](#platform-introduction)
* [Project Summary](#project-summary)
* [Participation](#participation)
* [Installation](#installation)
* [Development Environment](#development-environment)
* [Demo](#demo)
* [Todo List](#todo-list)

---

## Platform Introduction

EliteQuant is an open source forever free unified quant trading platform built by quant traders, for quant traders. It is dual listed on both [github](https://github.com/EliteQuant) and [gitee](https://gitee.com/EliteQuant).

The word unified carries two features.
- First it’s unified across backtesting and live trading. Just switch the data source to play with real money.
- Second it’s consistent across platforms written in their native langugages. It becomes easy to communicate with peer traders on strategies, ideas, and replicate performances, sparing language details.

Related projects include
- [A list of online resources on quantitative modeling, trading, and investment](https://github.com/EliteQuant/EliteQuant)
- [C++](https://github.com/EliteQuant/EliteQuant_Cpp)
- [Python](https://github.com/EliteQuant/EliteQuant_Python)
- [Matlab](https://github.com/EliteQuant/EliteQuant_Matlab)
- [R]()
- [C#]()
- [Excel](https://github.com/EliteQuant/EliteQuant_Excel)
- [Java]()
- [Scala]()
- [Go]()
- [Julia]()


## Project Summary

EliteQuant Excel is an Excel Add-in tool for pricing, portfolio and risk management. It uses QuantLib as the pricing engine for interest rate products, CDS, equities, and commodities. Simulation engine extends QuantLib into applications such as PFEs and CVAs.

For some more details, check out the [introductory blog post and youtube video](http://www.elitequant.com/2017/10/22/elitequant-excel-one/).

## Participation

Please feel free to report issues, fork the branch, and create pull requests. Any kind of contributions are welcomed and appreciated. Through shared code architecture, it also helps traders using other languges.

## Installation

No installation is needed, it's ready for use out of box. Just download compiled.zip and enjoy. 

#### Run Compiled

Download and upzip the file Compiled.zip located in the project root directory. Start Excel and open EliteQuantExcel-addin-x86.xll or EliteQuantExcel-addin-x64.dll from where you unzipped the file. A Ribbon tab called EliteQuant should show up. All functions start with eq, for example, call Excel function eqtimetoday() it will return today's date. There are demos such as Black-Scholes and historical market data workbooks in the workbooks dropdown menu.

#### Run Source Code

(1) Put boost, compiled QuantLib, and swigwin in folder D:\workspace. If you use different path, you have to change project reference path accordingly.

(2) Download and build the solution.

(3) Open the xll file from Excel and the Ribbon will show up.

## Development Environment

Below is the environment we are using
* boost c++ library v1.6.5
* QuantLib c++ library v1.11
* Swig Windows 3.0.12
* Visual studio community edition 2017
* Microsoft .Net framework 4.7 (you may need to downgrade it to meet that on your machine)
* Microsoft Office Excel Professional Plus 2016

## Demo

Black-Scholes option pricing

![Black Scholes](/resource/black_scholes.gif?raw=true "Black Scholes")

Historical market data

![Historical Data](/resource/hist_data.gif?raw=true "Historical Data")

## Todo List