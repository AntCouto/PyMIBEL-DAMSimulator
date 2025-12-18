#  PyMIBELSim-DAM: MIBEL Day-Ahead Market Simplified Simulator

This repository contains a simplified day-ahead electricity market clearing model for the Iberian Electricity Market (MIBEL), implemented using PyPSA. The model simulates hourly day-ahead market clearing based on linear programming and explicit supply and demand bids.

The model performs hourly clearing, explicitly represents the PT–ES interconnection capacity, and outputs zonal prices, traded energy volumes, and interconnection flows. Its objective is to reproduce the core economic clearing logic of the MIBEL day-ahead market while maintaining computational transparency and simplicity.

To achieve this, the model applies several simplifying assumptions, including:
- Each hour is cleared independently (no intertemporal constraints);
- The market is represented using two price zones only (PT and ES);
- Block bids, complex bids, and complex order conditions (e.g. minimum income, load gradient, linked bids) are not considered

Methodological details how the bids were constructed can be found in TradeRES – Performance Assessment of Markets, Deliverable D5.3 (Edition 2), specifically in the MIBEL case study (Iberian electricity market).  Reference: [1]	Estanqueiro, A., Couto, A., Algarvio, H., Sperber, E., Santos, G., Strbac, G., Sanchez Jimenez, I., Kochems, J., Sijm, J., Nienhaus, K., De Vries, L., Wang, N., Chrysanthopoulos, N., Martin Gregorio, N., Carvalho, R., Faia, R., & Vale, Z. (2024). D5.3 - Performance assessment of current and new market designs and trading mechanisms for national and regional markets (Ed. 2). TradeRES Project Deliverable Available at: https://traderes.eu/wp-content/uploads/2024/12/D5.3_TradeRES_Performance-assessment_Ed2.pdf 

# Requirements 
 - Python >= 3.9
 - pandas
 - numpy
 - pypsa
 - matplotlib
 - openpyxl
 - xlsxwriter
 - glpk (solver)


# How the Model Works 
-  Create a Execel file with hourly bid and interconnection data for 24 periods
- Separates PT/ES and BUY/SELL bids
- Builds a PyPSA network per hour
- Applies interconnection constraints
- Clears the market using linear optimisation
- Stores:
  - Zonal prices (PT & ES)
  - Interconnection flows
  - Traded energy per unit
-Exports  results to Excel 
- Plots 24h PT vs ES price comparison

# Running the Simulation
python MIBEL_DAM_v1-2.py

# Project Context and Status
This work has been developed at LNEG within the scope of the project “Man0EUvRE – Energy System Modelling for the Transition to Net-Zero 2050 for the EU via REPowerEU”, funded by CETPartnership, the European Partnership under the Joint Call 2022 for research proposals, co-funded by the European Commission (Grant Agreement No. 101069750), and with the funding organisations listed on the CETPartnership website.
