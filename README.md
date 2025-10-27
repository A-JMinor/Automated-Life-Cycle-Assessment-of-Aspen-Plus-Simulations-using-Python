# Automated-Life-Cycle-Assessment-from-Aspen-Plus-using-Python
This repository (casestudy2.ipynb) provides a **hands-on, end-to-end example** of how to perform a **transparent, parametric, and uncertainty-aware Life Cycle Assessment (LCA)** of a chemical process simulated in **Aspen Plus**, automated entirely from **Python**.

The tutorial case focuses on ε-caprolactam (CL) purification from water (functional unit: *1 kg purified CL*), comprising:
- one distillation column with reflux drum, condenser, and kettle reboiler,
- one evaporator,
- one pump, and
- one heat exchanger.


The workflow demonstrates how to:
1. Pull stream and utility data directly from Aspen Plus, 
2. Modify process parameters such as feed water amount from Python and re-run Aspen Plus 
3. Build a structured foreground inventory in Python, 
4. Define discrete and continuous parameters for scenario and sensitivity analyses, and 
5. Run the LCA using **Brightway2** and **lca-algebraic**, including parameter analysis, discrete/continuous choices, and Monte Carlo uncertainty propagation.

This codebase provides a compact, reproducible exemplar for others to learn, adapt, and extend toward their own process assessments.

---

### Key Features and Novelties

#### 1. Fully Automated Python ⇄ Aspen Plus Bridge
- Establishes a connection  
- Extracts detailed unit-operation data:
  - Mass and energy flows  
  - Utility duties (reboilers, condensers, pumps, exchangers)  
  - Stream compositions, temperatures, and pressures  
- Modifies parameters and re-runs Aspen Plus with updated inventories extracted for the LCA 
- Enables fully automated parametric regeneration of foreground inventories for process variants.


---

#### 2. Transparent Foreground–Background Coupling
- Merges Aspen-extracted inventories with **ecoinvent v3.9.1** background data.  
- Uses **Brightway2** and **lca-algebraic** for open-source LCA computation.  
- ISO 14040/44 compliance and applies the **EF 3.1** impact method.

---

#### 3. Structured LCA Workflow
The notebook walks through all LCA stages step by step.

---

#### 4. Parameterized Sensitivity Analysis
Includes both discrete and continuous parameters to illustrate uncertainty handling:

- **Discrete parameter:**  
  - Electricity source → switch between *German grid mix* and *wind power* datasets  

- **Continuous parameters:**  
  - Plant operating years  
  - Capacity factor 

All parameters are processed via lca-algebraic’s symbolic engine, allowing automated sensitivity and Monte-Carlo analyses.

---

#### 5. Modular Scenario Exploration
- Key assumptions (electricity mix, years, capacity factor, process water feed) can be modified directly in Python to generate new inventories and LCA results.
- Scenario loop varies the Aspen Plus water feed flowrate, triggering:
  - parameter back-substitution in Aspen,
  - re-simulation of process behavior,
  - regeneration of all stream and utility demands,
  - full recalculation of environmental impacts.
