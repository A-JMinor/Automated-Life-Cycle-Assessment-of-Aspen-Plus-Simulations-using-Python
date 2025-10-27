# -*- coding: utf-8 -*-
"""
Created on Mon Jul 25 17:19:09 2022

@author: Ann-Joelle
"""

import os                          # Import operating system interface
import win32com.client as win32    # Import COM
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import openpyxl
from openpyxl import workbook
from openpyxl import load_workbook
import math

#Unit calculations
m_to_inch = 39.3701
m_to_ft = 3.28084
m2_to_ft2 = 10.7639
m3_to_ft3 = 35.3147
kPa_to_psig = 0.145038
kPa_to_Torr = 7.50062
kg_to_lb = 2.20462
W_to_Btu_hr = 3.41
gal_to_m3 = 0.00378541
BTU_to_J = 1055.06
HP_to_W = 745.7
m3_s_to_gpm = 15850.3
kg_m3_to_lb_gal = 0.0083454
t_to_kg = 1000
HP_per_1000_Gal = 10
N_m2_to_psig = 0.000145038

#Operating year
hr_per_day = 24
day_per_year = 365
operating_factor = 0.9 

#cost factors of equations from Seider et al. 2006 
cost_index_2006 = 500


def distillationDWSTU_geometry(Application,
    nameDWSTU: str,
    name_inputstream_DWSTU: str,
    name_distallestream_DWSTU: str,
    tray_Spacing: float = 0.5,   # m
    top_space: float = 1.2,      # m
    bottom_space: float = 1.8,   # m
    d_rho: float = 0.288,        # lb/in^3 (material density for weight calc)
    d_diameter = None,
    flooding: float = 0.8        # fractional approach to flooding
):
    
    """
    Compute geometry/mechanical properties of the DWSTU column only.

    Parameters
    ----------
    Application : COM Aspen object
        Aspen Plus document (ProgID "Apwn.Document").
    nameDWSTU : str
        DWSTU block name (e.g., "DIST1").
    name_inputstream_DWSTU : str
        Feed stream name (e.g., "DST1IN").
    name_distallestream_DWSTU : str
        Distillate stream name (e.g., "DIST1TOP").
    tray_Spacing : float, default 0.5
        Tray spacing [m].
    top_space : float, default 1.2
        Top freeboard [m].
    bottom_space : float, default 1.8
        Bottom/kettle freeboard [m].
    d_rho : float, default 0.288
        Material density for weight calculation [lb/in³].
    d_diameter : float | None, optional
        Fixed diameter [m]. If None, diameter is estimated via Turton.
    flooding : float, default 0.8
        Fractional approach to flooding (multiplies the flooding velocity).

    Returns
    -------
    d_diameter : float
        Column diameter [m].
    d_volume : float
        Cylindrical volume between tan–tan [m³].
    d_height : float
        Total height incl. top/bottom spaces [m].
    d_weight_kg : float
        Structural weight [kg].
    no_of_trays : float
        Number of trays (Aspen gives decimal → rounded; may be adjusted if H/D < 3).
    d_tangent_tangent_length : float
        Tan–tan length [m].

    
    Assumptions & limitations
    -------------------------
    - ΔT is approximated as (T_hot_utility − T_bottom)
    - Default U = 1140 W/m²·°C and fouling = 0.9 are generic service values.
      Provide service-specific values if available.
    - Auto-selection of hot utility is based solely on bottom temperature and
      fixed utility setpoints; user-supplied temperature overrides this logic.
    - Area correlation internally uses ft² and converts back to m²; units are
      handled internally.
      
    Notes
    -----
    - Aspen 'OPT_RDV' (VAPOR/LIQUID) is toggled to retrieve both densities for the
      Turton correlation; Engine.Run2() is called after each toggle.
    - If H/D < 3, height is adjusted to 3·round(D) and no_of_trays recomputed.
    - Pressure/temperature bounds handled as in Seider-based code; warnings preserved.
    """
    
    # Trays / height
    no_of_trays = round(
        Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\ACT_STAGES").Value, 0
    )
    d_height = no_of_trays * tray_Spacing + top_space + bottom_space
    d_tangent_tangent_length = no_of_trays * tray_Spacing

    # Flows for Turton sizing (top section, worst case)
    d_RR = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\ACT_REFLUX").Value
    d_feed = Application.Tree.FindNode("\\Data\\Streams\\" + name_inputstream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value
    d_distillate = Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value
    d_liquidflow = d_RR * d_distillate
    d_vaporflow = d_liquidflow + d_distillate
    d_liquidflow_adapt = d_liquidflow
    d_vaporflow_adapt = d_vaporflow + d_feed

    # Densities from same stream by toggling OPT_RDV
    if Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\OPT_RDV").Value == 'VAPOR':
        d_vapor_rho = Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value / Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\VOLFLMX\\MIXED").Value
        Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\OPT_RDV").Value = 'LIQUID'
        Application.Engine.Run2()
        d_liquid_rho = Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value / Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\VOLFLMX\\MIXED").Value
        Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\OPT_RDV").Value = 'VAPOR'
        Application.Engine.Run2()
    else:  # LIQUID
        d_liquid_rho = Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value / Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\VOLFLMX\\MIXED").Value
        Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\OPT_RDV").Value = 'VAPOR'
        Application.Engine.Run2()
        d_vapor_rho = Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value / Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\VOLFLMX\\MIXED").Value
        Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\OPT_RDV").Value = 'LIQUID'
        Application.Engine.Run2()

    # Turton capacity / flooding
    d_flow_param = (d_liquidflow_adapt / d_vaporflow_adapt) * (d_vapor_rho / d_liquid_rho) ** 0.5
    d_capacity_param = 10 ** (-1.0262 - 0.63513 * np.log10(d_flow_param) - 0.20097 * (np.log10(d_flow_param)) ** 2)
    d_flooding_velocity = d_capacity_param * 1.3 * ((d_liquid_rho - d_vapor_rho) / d_vapor_rho) ** 0.5  # ft/s
    d_vapor_velocity = (d_flooding_velocity / m_to_ft) * flooding  # m/s

    d_active_area = d_vaporflow / (d_vapor_rho * d_vapor_velocity)
    if d_diameter is None:
        d_diameter = (4 * d_active_area / np.pi) ** 0.5

    # H/D ≥ 3
    if d_height / round(d_diameter, 0) < 3:
        d_height = 3 * round(d_diameter, 0)
        no_of_trays = (d_height - top_space - bottom_space) / tray_Spacing
        d_tangent_tangent_length = no_of_trays * tray_Spacing

    if d_height > 60:
        print("Warning distillation tower too high")

    d_volume = (np.pi / 4) * d_diameter ** 2 * d_height

    # Seider wall thickness + weight (needs pressure & temp)
    d_lowest_pressure = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\PTOP").Value  # kPa
    if d_lowest_pressure <= 34.5:
        d_design_pressure = 10.0  # psig
    elif d_lowest_pressure <= 6895:
        d_design_pressure = np.exp(0.60608 + 0.91615 * np.log(d_lowest_pressure * kPa_to_psig) +
                                   0.0015655 * (np.log(d_lowest_pressure * kPa_to_psig)) ** 2)
    else:
        print("Warning!! operating pressure too high for cost calculation of distillation column! Look for another cost equation")

    d_highest_temp = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\BOTTOM_TEMP").Value
    d_design_temp = (d_highest_temp - 273.15) * 9 / 5 + 32.0 + 50.0

    if d_design_temp < 200.0:
        d_E_modulus = 30.2e6
    elif d_design_temp < 400.0:
        d_E_modulus = 29.5e6
    elif d_design_temp < 650.0:
        d_E_modulus = 28.3e6
    else:
        d_E_modulus = 26.0e6

    if d_design_temp <= 750:
        d_allowable_stress = 15000
    elif d_design_temp <= 800:
        d_allowable_stress = 14750
    elif d_design_temp <= 850:
        d_allowable_stress = 14200
    elif d_design_temp <= 900:
        d_allowable_stress = 13100
    else:
        print("Warning!! distillation design temperature is too high for wall thickness calculation from seider ")

    d_tE1 = 0.25
    if d_lowest_pressure >= 101:
        while True:
            d_tE = 0.22 * (((d_diameter * m_to_inch) + d_tE1) + 18) * (d_tangent_tangent_length * m_to_inch) ** 2 / (d_allowable_stress * ((d_diameter * m_to_inch) + d_tE1) ** 2)
            if abs((d_tE1 - d_tE) / d_tE1) <= 1e-3:
                break
            d_tE1 = d_tE
        d_tp = (d_design_pressure * (d_diameter * m_to_inch)) / (2 * d_allowable_stress - 1.2 * d_design_pressure)
        d_t_total = (d_tp + d_tE + d_tp) / 2
    else:
        while True:
            d_tE = 1.3 * ((d_diameter * m_to_inch) + d_tE1) * ((d_design_pressure * (d_tangent_tangent_length * m_to_inch)) / (d_E_modulus * ((d_diameter * m_to_inch) + d_tE1))) ** 0.4
            if abs((d_tE1 - d_tE) / d_tE1) <= 1e-3:
                break
            d_tE1 = d_tE
        check = d_tE / (d_diameter * m_to_inch)
        if check >= 0.05:
            print("Warning!! the wall thickness of the distillation column did not pass the methods.")
        d_tEC = (d_tangent_tangent_length * m_to_inch) * (0.18 * (d_diameter * m_to_inch) - 2.2) * 1e-5 - 0.19
        d_t_total = d_tE + 0.125 + (d_tEC if d_tEC > 0 else 0)

    if d_t_total < 0.25:
        d_t_total = 0.25

    d_weight_lb = np.pi * d_t_total * d_rho * ((d_diameter * m_to_inch) + d_t_total) * ((d_tangent_tangent_length * m_to_inch) + 0.8 * (d_diameter * m_to_inch))
    d_weight_kg = d_weight_lb / kg_to_lb

    return d_diameter, d_volume, d_height, d_weight_kg, no_of_trays, d_tangent_tangent_length





def refluxdrumDWSTU_geometry(
    Application,
    nameDWSTU: str,
    name_distallestream_DWSTU: str,
    drum_residence_time: float = 300.0,  # s
    drum_filled: float = 0.5,            # –
    drum_l_to_d: float = 3.0,            # –
    r_rho: float = 0.288                 # lb/in^3
):
    """
    Geometry/mechanical sizing of the reflux drum.

    Requirements (Aspen Plus)
    -------------------------
    - DWSTU block `nameDWSTU` exists and is converged.
    - Distillate stream `name_distallestream_DWSTU` exists.

    Parameters
    ----------
    drum_residence_time : float, default 300.0
        Liquid holdup time [s].
    drum_filled : float, default 0.5
        Fraction of drum volume occupied by liquid (–).
    drum_l_to_d : float, default 3.0
        Length-to-diameter ratio (–).
    r_rho : float, default 0.288
        Vessel material density for weight calc [lb/in³].

    Returns
    -------
    drum_volume : float
        Working volume [m³].
    drum_diameter : float
        Diameter [m].
    drum_length : float
        Tan–tan length [m].
    drum_weight_kg : float
        Estimated shell weight [kg].

    Notes/Assumptions
    -----------------
    - Horizontal cylindrical drum.
    - Holdup basis: volumetric flow from distillate rate and reflux (Aspen tags as in code).
    - Wall thickness per Seider correlations with a 0.25 in minimum.
    - Units: Aspen reads in SI; density input `r_rho` is in lb/in³ by convention here.
    """
    
    #  geometry and mechanics 
    d_distillate = Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value
    d_RR = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\ACT_REFLUX").Value
    d_liquid_rho = Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\MASSFLMX\\MIXED").Value / Application.Tree.FindNode("\\Data\\Streams\\" + name_distallestream_DWSTU + "\\Output\\STR_MAIN\\VOLFLMX\\MIXED").Value
    d_lowest_pressure = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\PTOP").Value  # kPa

    drum_liquid_flowrate  = d_distillate * (1 + d_RR)        # kg/s
    drum_liquid_density = d_liquid_rho                       # kg/m3
    drum_flowrate = drum_liquid_flowrate / drum_liquid_density        # m3/s
    drum_hold_up = drum_residence_time * drum_flowrate
    drum_volume = drum_hold_up / drum_filled
    drum_diameter = (drum_volume*(4/np.pi))**(1/3)
    drum_length = drum_l_to_d * drum_diameter

    drum_lowest_pressure = d_lowest_pressure

    if drum_lowest_pressure <= 34.5: # kPa
        drum_design_pressure = 10.0 # psig
    elif drum_lowest_pressure > 34.5 and drum_lowest_pressure <= 6895:
        drum_design_pressure = np.exp(
            0.60608 +
            0.91615 * np.log(drum_lowest_pressure*kPa_to_psig) +
            0.0015655*np.log(drum_lowest_pressure*kPa_to_psig)**2
        )
    else:
        print("Warning!! operating pressure too high for cost calculation of reflux drum! Look for another cost equation")

    drum_temp = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\BOTTOM_TEMP").Value  # K
    drum_design_temp = (drum_temp - 273.15) * 9/5 + 32.0 + 50.0  # °F

    if drum_design_temp >= -20.0 and drum_design_temp < 200.0 :
        drum_E_modulus = 30.2 *10**6
    elif drum_design_temp >= -200.0 and drum_design_temp < 400.0 :
        drum_E_modulus = 29.5 *10**6
    elif drum_design_temp >= -400.0 and drum_design_temp < 650.0 :
        drum_E_modulus = 28.3 *10**6
    elif drum_design_temp >= -650.0 and drum_design_temp < 700.0 :
        drum_E_modulus = 26.0 *10**6
    else:
        print("Warning!! Design temperature is too high for carbon steel, use another material")

    if drum_design_temp >= -20 and drum_design_temp <= 750:
        drum_allowable_stress = 15000
    elif drum_design_temp <= 800:
        drum_allowable_stress = 14750
    elif drum_design_temp <= 850:
        drum_allowable_stress = 14200
    elif drum_design_temp <= 900:
        drum_allowable_stress = 13100
    else:
        print("Warning!! reactor design temperature is too high for wall thickness calculation from seider ")

    if drum_lowest_pressure * 0.001 >= 101:         # 101 kPa ~ atm
        drum_t_total = (drum_design_pressure * (drum_diameter*m_to_inch))/(2*drum_allowable_stress-1.2*drum_design_pressure)
    else:   # vacuum operation
        drum_tE1 = 0.25
        error = 1
        while abs(error) > 0.001:
            drum_tE = 1.3 * ((drum_diameter*m_to_inch)+drum_tE1) * ((drum_design_pressure * (drum_length*m_to_inch)) /(drum_E_modulus*((drum_diameter*m_to_inch)+drum_tE1)))**0.4
            error = (drum_tE1 - drum_tE)/drum_tE1
            drum_tE1 = drum_tE

        check = drum_tE/(drum_diameter*m_to_inch)
        if check >= 0.05 :
            print("Warning!! the wall thickness of the distillation column did not pass the methods. ")

        drum_tEC = (drum_length*m_to_inch) * (0.18*(drum_diameter*m_to_inch)-2.2)*10**(-5) - 0.19

        if drum_tEC > 0 :
            drum_t_total = drum_tEC + drum_tE + 0.125       # 0.125 corrosion allowance
        else :
            drum_t_total = drum_tE + 0.125

    if drum_t_total < 0.25:        # minimum wall thickness
        drum_t_total = 0.25

    drum_weight = np.pi*drum_t_total * r_rho * ((drum_diameter*m_to_inch) + drum_t_total) * ((drum_length*m_to_inch) + 0.8 * (drum_diameter*m_to_inch))

    drum_weight_kg = drum_weight / kg_to_lb
    return drum_volume, drum_diameter, drum_length, drum_weight_kg




def kettleDWSTU_geometry(Application, nameDWSTU, kettle_hotutility_temperature=None, kettle_U=None, fouling_factor=None):
    """
    Size the DWSTU kettle reboiler (geometry only). No costing.

    Aspen requirements
    ------------------
    A DWSTU block named `nameDWSTU` with:
      - \\Data\\Blocks\\{nameDWSTU}\\Output\\REB_DUTY        (W)
      - \\Data\\Blocks\\{nameDWSTU}\\Output\\BOTTOM_TEMP     (K)
      - \\Data\\Blocks\\{nameDWSTU}\\Input\\PBOT             (kPa)

    Behavior & defaults
    -------------------
    - If `kettle_hotutility_temperature` is None, pick utility by BOTTOM_TEMP:
        <=120°C→138.9°C, <=170°C→186°C, <=255°C→270°C, <=300°C→337.8°C, <=380°C→400°C (all in K internally).
      If above range, use T_bottom + 30 K to keep ΔT > 0.
    - If `kettle_U` is None, use 1140 W/m²·°C.
    - If `fouling_factor` is None, use 0.9.
    - Area uses the original simple ΔT = (T_hot − T_bottom) (not full LMTD).

    Parameters
    ----------
    Application : Aspen COM object
    nameDWSTU : str
        DWSTU block name in Aspen.
    kettle_hotutility_temperature : float | None, optional (K)
    kettle_U : float | None, optional (W/m²·°C)
    fouling_factor : float | None, optional (>0)

    Returns
    -------
    kettle_Q : float
        Reboiler duty [W].
    kettle_area_m2 : float
        Required heat-transfer area [m²].
    kettle_pressure_psig : float
        Bottom pressure [psig] (PBOT converted from kPa).

    Raises
    ------
    ValueError
        If `fouling_factor` ≤ 0 or ΔT ≤ 0.
    """


    # Aspen reads
    kettle_Q = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\REB_DUTY").Value
    kettle_cold = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\BOTTOM_TEMP").Value  # [K]

    # Auto-select hot utility if not provided
    if kettle_hotutility_temperature is None:
        T = kettle_cold  # K
        if T <= (120 + 273.15):
            kettle_hotutility_temperature = 138.9 + 273.15  # LP Steam
        elif T <= (170 + 273.15):
            kettle_hotutility_temperature = 186.0 + 273.15  # MP Steam
        elif T <= (255 + 273.15):
            kettle_hotutility_temperature = 270.0 + 273.15  # HP Steam
        elif T <= (300 + 273.15):
            kettle_hotutility_temperature = 337.8 + 273.15  # Fuel Oil No. 2
        elif T <= (380 + 273.15):
            kettle_hotutility_temperature = 400.0 + 273.15  # Dowtherm A
        else:
            print("kettle temperature out of range")
            kettle_hotutility_temperature = T + 30.0  # fallback to keep LMTD > 0

    # Defaults for U and fouling
    if kettle_U is None:
        kettle_U = 1140  # W/m²·°C
    if fouling_factor is None:
        fouling_factor = 0.9
    if fouling_factor <= 0:
        raise ValueError("fouling_factor must be > 0")

    # LMTD surrogate (simple ΔT per original)
    kettle_LMTD = kettle_hotutility_temperature - kettle_cold
    if kettle_LMTD <= 0:
        raise ValueError(f"LMTD <= 0: hot={kettle_hotutility_temperature} K, cold={kettle_cold} K")

    # Area in m² (compute via ft² correlation path, then convert back)
    kettle_area_ft2 = kettle_Q / (kettle_U * kettle_LMTD * fouling_factor) * m2_to_ft2
    kettle_area_m2 = kettle_area_ft2 / m2_to_ft2

    # Pressure for costing
    kettle_pressure_psig = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\PBOT").Value * kPa_to_psig

    return kettle_Q, kettle_area_m2, kettle_pressure_psig




def condenserDWSTU_geometry(
    Application,
    nameDWSTU: str,
    fouling_factor=None,
    cond_cold_in=None,   # [°C] if None → 30
    cond_cold_out=None,  # [°C] if None → 45
    cond_U=None          # [W/m²·°C] if None → 1140
):
    """
    Geometry/sizing for the DWSTU condenser (no costing).

    Summary
    -------
    Computes condenser heat duty (Q) and heat-transfer area using LMTD:
        area = Q / (U * LMTD * fouling).
    If cold-side temperatures or U are not provided, sensible defaults are used
    to preserve original behavior.

    Aspen requirements
    ------------------
    Uses these nodes on the DWSTU block:
      - \\Data\\Blocks\\{nameDWSTU}\\Output\\COND_DUTY     (W)
      - \\Data\\Blocks\\{nameDWSTU}\\Output\\DISTIL_TEMP   (K)
      - \\Data\\Blocks\\{nameDWSTU}\\Input\\PTOP           (kPa)

    Parameters
    ----------
    Application : COM Aspen object
        Aspen Plus document.
    nameDWSTU : str
        DWSTU block name (e.g., "DIST1").
    fouling_factor : float | None
        Multiplier on U·LMTD. If None → 0.9 (keeps consistency with other units).
    cond_cold_in : float | None
        Cooling-water inlet temperature [°C]. If None → 30.
    cond_cold_out : float | None
        Cooling-water outlet temperature [°C]. If None → 45.
    cond_U : float | None
        Overall heat-transfer coefficient [W/m²·°C]. If None → 1140.

    Returns
    -------
    cond_Q : float
        Condenser duty [W].
    cond_area_m2 : float
        Required heat-transfer area [m²].
    cond_pressure_psig : float
        Condenser/top pressure [psig] (from PTOP, converted kPa→psig).

    Assumptions (screening level)
    -----------------------------
    - Constant U and single LMTD based on bulk hot temperature (DISTIL_TEMP).
    - Default fouling = 0.9 if not provided.
    - If cond_cold_in == cond_cold_out, falls back to ΔT = (T_hot − T_cold_out).
    """
    # Defaults that preserve original behavior
    if cond_U is None:
        cond_U = 1140.0
    if fouling_factor is None:
        fouling_factor = 0.9
    if cond_cold_in is None:
        cond_cold_in = 30.0
    if cond_cold_out is None:
        cond_cold_out = 45.0

    # Aspen reads
    cond_Q = abs(Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\COND_DUTY").Value)  # [W]
    cond_hot = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Output\\DISTIL_TEMP").Value   # [K]

    # Convert cold-side to K
    Tci = cond_cold_in + 273.15
    Tco = cond_cold_out + 273.15

    # LMTD (guard for equal in/out temps)
    if Tco != Tci:
        cond_LMTD = (Tco - Tci) / np.log((cond_hot - Tci) / (cond_hot - Tco))
    else:
        cond_LMTD = cond_hot - Tco

    # Area via ft² correlation path (kept from original), then back to m²
    cond_area_ft2 = cond_Q / (cond_U * cond_LMTD * fouling_factor) * m2_to_ft2
    cond_area_m2 = cond_area_ft2 / m2_to_ft2

    # Pressure (for costing)
    cond_pressure_psig = Application.Tree.FindNode("\\Data\\Blocks\\" + nameDWSTU + "\\Input\\PTOP").Value * kPa_to_psig

    return cond_Q, cond_area_m2, cond_pressure_psig



def fallingEVAPORATORS_geometry(
    Application,
    No_Evaporators: int,
    souders_brown_param: float = 0.35,   # vertical vane type (Turton)
    L_D_ratio: float = 2.5,              # Turton
    evap_U: float = 850,                 # W/m²·°C
    evap_hotutility_temperature=None,    # list[K] or None; if None -> LP steam for all
    fouling_factor: float = 0.9
):
    """
    Geometry/sizing for falling-film evaporators.

    Aspen requirements
    ------------------
    - Streams: EVAP{i}TOP, EVAP{i}BOT (i = 1..No_Evaporators)
    - Blocks:  EVAP{i} with QCALC (duty) and TEMP (cold-side temperature)

    Parameters
    ----------
    No_Evaporators : int
        Number of evaporators (i = 1..N).
    souders_brown_param : float, default 0.35
        Souders–Brown K-value for vapor disengagement [m/s·sqrt(kg/m³)].
    L_D_ratio : float, default 2.5
        Shell length-to-diameter ratio (-).
    evap_U : float, default 850
        Overall heat-transfer coefficient [W/m²·°C].
    evap_hotutility_temperature : list[float] | None
        Hot-utility temps [K] per evaporator; if None, uses LP steam (139.9°C) for all.
    fouling_factor : float, default 0.9
        Multiplicative fouling factor (-).

    Returns
    -------
    evap_volume : ndarray (N,)
        Vessel volume [m³] for each evaporator.
    evap_Q : ndarray (N,)
        Duty [W] from \\Data\\Blocks\\EVAP{i}\\Output\\QCALC.
    evap_area_m2 : ndarray (N,)
        Required heat-transfer area [m²].
    evap_diameter : ndarray (N,) or float if N=1
        Shell diameter [m].
    evap_length : ndarray (N,) or float if N=1
        Shell length (tan–tan) [m].


    Notes
    -----
    - Area is computed via A = Q / (U · ΔT · fouling), with ΔT = Thot − Tcold.
    - Souders–Brown sizing: vmax = K·sqrt((ρL−ρV)/ρV); D from volumetric vapor rate.
    """
    
    # default hot-utility list if not provided
    if evap_hotutility_temperature is None:
        evap_hotutility_temperature = [139.9 + 273.15] * No_Evaporators  # LP steam for all

    evap_volume = np.zeros(No_Evaporators)
    evap_Q = np.zeros(No_Evaporators)
    evap_area_m2 = np.zeros(No_Evaporators)
    evap_diameter = np.zeros(No_Evaporators)
    evap_length = np.zeros(No_Evaporators)

    for i in range(1, No_Evaporators + 1):
        # Vapor disengagement diameter (Souders–Brown)
        evap_vapor_flowrate = Application.Tree.FindNode("\\Data\\Streams\\EVAP{}TOP\\Output\\STR_MAIN\\VOLFLMX\\MIXED".format(i)).Value  # m³/s
        evap_rho_vapor = Application.Tree.FindNode("\\Data\\Streams\\EVAP{}TOP\\Output\\STR_MAIN\\MASSFLMX\\MIXED".format(i)).Value / Application.Tree.FindNode("\\Data\\Streams\\EVAP{}TOP\\Output\\STR_MAIN\\VOLFLMX\\MIXED".format(i)).Value  # kg/m³
        evap_rho_liquid = Application.Tree.FindNode("\\Data\\Streams\\EVAP{}BOT\\Output\\STR_MAIN\\MASSFLMX\\MIXED".format(i)).Value / Application.Tree.FindNode("\\Data\\Streams\\EVAP{}BOT\\Output\\STR_MAIN\\VOLFLMX\\MIXED".format(i)).Value  # kg/m³

        vmax = souders_brown_param * math.sqrt((evap_rho_liquid - evap_rho_vapor) / evap_rho_vapor)  # m/s
        evap_diameter[i-1] = math.sqrt(evap_vapor_flowrate * 4.0 / (math.pi * vmax))
        evap_length[i-1] = L_D_ratio * evap_diameter[i-1]
        evap_volume[i-1] = math.pi / 4.0 * evap_diameter[i-1] ** 2 * evap_length[i-1]

        # Area (ΔT as simple Thot − Tcold, consistent with your original)
        evap_Q[i-1] = Application.Tree.FindNode("\\Data\\Blocks\\EVAP{}\\Output\\QCALC".format(i)).Value  # W
        evap_cold = Application.Tree.FindNode("\\Data\\Blocks\\EVAP{}\\Input\\TEMP".format(i)).Value      # K
        evap_LMTD = evap_hotutility_temperature[i-1] - evap_cold
        evap_area_ft2 = evap_Q[i-1] / (evap_U * evap_LMTD * fouling_factor) * m2_to_ft2
        evap_area_m2[i-1] = evap_area_ft2 / m2_to_ft2

    # keep your original scalar-collapse behavior for N=1
    if No_Evaporators == 1:
        evap_diameter = float(np.sum(evap_diameter))
        evap_length = float(np.sum(evap_length))

    return evap_volume, evap_Q, evap_area_m2, evap_diameter, evap_length



def heatexchanger_geometry(
    Application,
    No_Heat_Exchanger: int,
    fouling_factor: float = 0.9,
    F_M_a: float = 1.75,
    F_M_b: float = 0.13,
):
    """
    Geometry/sizing for heat exchangers.

    Aspen requirements
    ------------------
    Blocks named E01, E02, ... with Aspen outputs as in the original code:
      - For standard HX: HX_DUTY, U, HX_DTLM
      - For refrigeration (< 40°C cold side): QCALC, B_TEMP
      - For fired heaters (> 252°C): QCALC (area not computed)

    Behavior
    --------------------------------
    This function auto-detects the equipment type from the Aspen block and its
    temperatures, then applies the matching sizing/costing path:
    
    1) Standard heat exchanger (block exposes MODE and HX variables):
       - Used when the block is a regular HX (neither refrigeration nor fired heater).
       - Area is computed from Aspen’s own values:
           A_ft² = HX_DUTY / (U * HX_DTLM * fouling_factor)
    
    2) Refrigeration cooler (very cold service):
       - Triggered when the relevant block-side temperature E_T < 40 °C.
       - Uses a fixed refrigerant approach (ammonia):
           U = 850 W/m²·°C, T_refrigerant = –13 °C
       - Area from: A_ft² = Q / (U * (T_process – T_refrigerant) * fouling_factor)
    
    3) Fired heater (fuel-fired):
       - Triggered when E_T is between 252 °C and 300 °C.
       - Treated as a fired heater; area is not used (as in the original code).
       - Cost is based on heat duty (BTU/h) correlations.
    
    4) Dowtherm A heater (very hot service):
       - Triggered when E_T > 300 °C.
       - Treated as a special fired heater for Dowtherm A; duty-based cost.
    
    Additional notes:
    - If the computed/estimated area < 150 ft² → the “double-pipe” path is used.
    - If area ≥ 150 ft² → the “shell-and-tube” path is used, and the tube-length
      correction factor E_FL (default 1.05) is applied in the cost correlation.
    - Water/steam services default to carbon steel (FM = 1); otherwise the Seider
      FM correlation is used.

    Parameters
    ----------
    No_Heat_Exchanger : int
        Number of exchangers (i = 1..N).
    fouling_factor : float, default 0.9
        Multiplicative fouling factor (-).

    Returns
    -------
    E_Q : ndarray (N,)
        Duty [W] (signs preserved as in original branches).
    E_area_m2 : ndarray (N,)
        Heat-transfer area [m²] (0 for fired heaters, identical to original).
    E_T : ndarray (N,)
        Block temperature used for branching [K].
    a : ndarray (N,)               1 if standard HX (Input\\MODE present), else 0
    E_FM : ndarray (N,)
        Material factor computed per branch (see Behavior).
    """
    
    E_T = np.zeros(No_Heat_Exchanger)
    E_Q = np.zeros(No_Heat_Exchanger)
    E_area_m2 = np.zeros(No_Heat_Exchanger)
    E_FM = np.zeros(No_Heat_Exchanger)

    for i in range(1, No_Heat_Exchanger + 1):
        try:
            E_T[i - 1] = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\COLD_TEMP".format(i)).Value
        except:
            pass
        try:
            E_T[i - 1] = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\B_TEMP".format(i)).Value
        except:
            pass

        is_standard = False
        try:
            Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Input\\MODE".format(i)).Value
            is_standard = True
        except:
            is_standard = False

        if is_standard:
            E_Q[i - 1] = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\HX_DUTY".format(i)).Value
            E_U = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Input\\U".format(i)).Value
            E_LMTD = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\HX_DTLM".format(i)).Value
            E_area_ft2 = E_Q[i - 1] / (E_U * E_LMTD * fouling_factor) * m2_to_ft2
            E_area_m2[i - 1] = E_area_ft2 / m2_to_ft2

            watercontent = None
            try:
                streamname = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\COLDIN".format(i)).Value
                watercontent = Application.Tree.FindNode(
                    "\\Data\\Streams\\{}\\Output\\STR_MAIN\\MASSFRAC\\MIXED\\WATER".format(streamname)
                ).Value
            except:
                pass
            try:
                streamname = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\HOTIN".format(i)).Value
                watercontent = Application.Tree.FindNode(
                    "\\Data\\Streams\\{}\\Output\\STR_MAIN\\MASSFRAC\\MIXED\\WATER".format(streamname)
                ).Value
            except:
                pass

            if (watercontent is not None) and (watercontent >= 0.999):
                E_FM[i - 1] = 1.0
            else:
                E_FM[i - 1] = F_M_a + (E_area_ft2 / 100.0) ** F_M_b

        elif (E_T[i - 1] > 252 + 273.15) and (E_T[i - 1] <= 300 + 273.15):
            E_Q[i - 1] = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\QCALC".format(i)).Value
            E_area_m2[i - 1] = 0.0
            E_FM[i - 1] = 1.7

        elif E_T[i - 1] > 300 + 273.15:
            E_Q[i - 1] = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\QCALC".format(i)).Value
            E_area_m2[i - 1] = 0.0
            E_FM[i - 1] = 1.7

        elif E_T[i - 1] < 40 + 273.15:
            E_Q[i - 1] = -Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\QCALC".format(i)).Value
            E_U = 850.0
            refrig_T = -13.0 + 273.15
            E_LMTD = Application.Tree.FindNode("\\Data\\Blocks\\E0{}\\Output\\B_TEMP".format(i)).Value - refrig_T
            E_area_ft2 = E_Q[i - 1] / (E_U * E_LMTD * fouling_factor) * m2_to_ft2
            E_area_m2[i - 1] = E_area_ft2 / m2_to_ft2
            E_FM[i - 1] = 1.75 + (E_area_ft2 / 100.0) ** 0.13

    return E_Q, E_area_m2, E_T, E_FM




def estimate_tube_number(heat_transfer_area, tube_outer_diameter, length, n, min_tubes=20):
    """
    Estimate the number of tubes required to achieve a specified heat transfer area with a given tube diameter and a fixed length.

    Parameters:
    - heat_transfer_area (float): The total required heat transfer area (m^2).
    - tube_outer_diameter (float): The outer diameter of the tubes (m).
    - length (float): The fixed length of the tubes (m).
    - n (int): Maximum allowable number of tubes.
    - min_tubes (int): Minimum number of tubes typically needed for shell and tube heat exchangers, default is 20.

    Returns:
    - int: The estimated total number of tubes, adjusted for minimum practical limits unless only one tube is required.
    """
    # Calculate the surface area of one tube
    tube_surface_area_per_meter = math.pi * tube_outer_diameter
    tube_surface_area = tube_surface_area_per_meter * length  # Total surface area for the given length

    # Calculate the number of tubes needed to achieve the desired total heat transfer area
    total = math.ceil(heat_transfer_area / tube_surface_area)

    # Adjust the tube count based on specific requirements
    if total > 1 and total < min_tubes:
        total = min_tubes  # Set to minimum tubes if below threshold but more than one
    if total > n:
        return 0  # Return 0 if the total exceeds the maximum allowable number of tubes

    return total


def calculate_baffle_spacing(shell_diameter, baffle_cut_percentage):
    """
    Estimate the number of baffles and their spacing according to Kern's method.

    Parameters:
    - shell_diameter (float): Internal diameter of the shell (m).
    - baffle_cut_percentage (int): Percentage of baffle cut (%).

    Returns:
    tuple: Number of baffles and recommended baffle spacing (m).
    """
    # Calculate number of baffles
    num_baffles = max(0, round(shell_diameter / 0.9))  # Example correlation, adjust as needed

    # Calculate baffle spacing based on cut percentage and number of baffles
    if num_baffles > 0:
        baffle_spacing = shell_diameter / (num_baffles + (baffle_cut_percentage / 100))
    else:
        baffle_spacing = 0  # Set baffle spacing to 0 if there are no baffles

    return num_baffles, baffle_spacing




def calculate_shell_diameter(num_tubes, tube_outer_diameter, pitch_type='t', min_shell_diameter=0.15):
    """
    Calculates the shell diameter for a shell and tube heat exchanger, enforcing a minimum shell diameter.

    Parameters:
    num_tubes (int): Number of tubes in the heat exchanger.
    tube_outer_diameter (float): Outer diameter of each tube (in meters).
    pitch_type (str): Type of pitch ('t' for triangular, 's' for square).
    min_shell_diameter (float): Minimum shell diameter (in meters), default is 0.15 meters.

    Returns:
    float: Shell diameter (in meters), adjusted for minimum size.
    """
    # Set the pitch factor based on the pitch type
    if pitch_type == 't':
        pitch_factor = 1.1  # Typical for triangular pitch
    elif pitch_type == 's':
        pitch_factor = 1.25  # Typical for square pitch
    else:
        raise ValueError("Invalid pitch type. Use 't' for triangular or 's' for square.")

    # Calculate the pitch (center-to-center distance between tubes)
    pitch = tube_outer_diameter * pitch_factor

    # Calculate the approximate diameter of the tube bundle
    tube_bundle_diameter = math.sqrt((num_tubes * pitch**2) / math.pi)

    # Estimate the shell diameter (usually 30% larger than the tube bundle diameter)
    shell_diameter = tube_bundle_diameter * 1.3

    # Apply the minimum shell diameter rule
    shell_diameter = max(shell_diameter, min_shell_diameter)

    return shell_diameter


def calculate_shelltubeexchanger_weight(shell_diameter, tube_length, tube_outer_diameter, num_tubes, baffle_spacing, shell_thickness=0.0127, tube_thickness=0.00211, baffle_thickness=0.00635, shell_steel_density=7850, tube_steel_density=7850):
    """
    Calculates the approximate weight of a shell and tube heat exchanger, including consideration for the presence of baffles based on the provided baffle spacing.

    Parameters:
    shell_diameter (float): Diameter of the shell in meters.
    tube_length (float): Length of the tubes in meters.
    tube_outer_diameter (float): Outer diameter of the tubes in meters.
    num_tubes (int): Number of tubes.
    baffle_spacing (float): Spacing between baffles in meters.
    shell_thickness (float): Thickness of the shell in meters (default is typical for carbon steel).
    tube_thickness (float): Thickness of the tubes in meters (default is typical for standard carbon steel tubes).
    baffle_thickness (float): Thickness of the baffles in meters (default is typical thickness).
    shell_steel_density (float): Density of the shell side steel in kg/m^3 (default is 7850 for carbon steel).
    tube_steel_density (float): Density of the shell side steel in kg/m^3 (default is 7850 for carbon steel).

    Returns:
    float: Total weight of the heat exchanger in kilograms.
    """
    # Calculate shell volume
    shell_inner_diameter = shell_diameter - 2 * shell_thickness
    shell_volume = math.pi * (shell_diameter**2 - shell_inner_diameter**2) / 4 * tube_length

    # Calculate tube volume
    tube_inner_diameter = tube_outer_diameter - 2 * tube_thickness
    tube_volume = num_tubes * math.pi * (tube_outer_diameter**2 - tube_inner_diameter**2) / 4 * tube_length

    # Initialize baffle weight
    baffle_weight = 0

    # Determine if baffles are needed
    if baffle_spacing > 0 and baffle_spacing < tube_length:
        # Calculate baffle volume
        baffle_count = math.floor(tube_length / baffle_spacing)
        baffle_area = shell_inner_diameter * shell_inner_diameter * math.pi / 4
        baffle_volume = baffle_area * baffle_thickness * baffle_count

        # Calculate baffle weight
        baffle_weight = baffle_volume * shell_steel_density

    # Calculate weights
    shell_weight = shell_volume * shell_steel_density  + baffle_weight
    tube_weight = tube_volume * tube_steel_density

    # Total weight
    total_weight = shell_weight + tube_weight + baffle_weight

    return total_weight, shell_weight, tube_weight


def pumps_geometry(Application, No_pumps):
    """
    Geometry/performance for pumps.

    Aspen requirements
    ------------------
    Blocks P01, P02, ... exposing:
      - Output\\VFLOW         [m3/s]
      - Output\\HEAD_CAL      [J/kg]
      - Output\\ELEC_POWER    [W]
      - Output\\BAL_MASI_TFL  [kg/s] (used to get density → BH)

    Returns (5-tuple; same units as before)
    ---------------------------------------
    pump_head : ndarray (N,)             Total head [ft]
    pump_flowrate : ndarray (N,)         Volumetric flow [m3/s]
    pump_size_factor : ndarray (N,)      (gpm) * (ft)^0.5
    pump_break_horsepower : ndarray (N,) Break horsepower [HP]
    pump_electricity_W : ndarray (N,)    Electrical power [W] from Aspen
    """
    
    earth_acceleration = 9.81  # m/s²

    pump_flowrate = np.zeros(No_pumps)
    pump_head_J_kg = np.zeros(No_pumps)
    pump_head = np.zeros(No_pumps)
    pump_size_factor = np.zeros(No_pumps)
    pump_break_horsepower = np.zeros(No_pumps)
    pump_electricity_W = np.zeros(No_pumps)

    for i in range(1, No_pumps + 1):
        # Aspen reads
        pump_flowrate[i-1] = Application.Tree.FindNode(f"\\Data\\Blocks\\P0{i}\\Output\\VFLOW").Value       # m3/s
        pump_head_J_kg[i-1] = Application.Tree.FindNode(f"\\Data\\Blocks\\P0{i}\\Output\\HEAD_CAL").Value    # J/kg
        pump_electricity_W[i-1] = Application.Tree.FindNode(f"\\Data\\Blocks\\P0{i}\\Output\\ELEC_POWER").Value  # W

        # Head [ft]
        pump_head[i-1] = pump_head_J_kg[i-1] / earth_acceleration * m_to_ft

        # Size factor (gpm * ft^0.5)
        pump_size_factor[i-1] = (pump_flowrate[i-1] * m3_s_to_gpm) * pump_head[i-1]**0.5

        # Density [lb/gal] for BH
        mass_flow_kg_s = Application.Tree.FindNode(f"\\Data\\Blocks\\P0{i}\\Output\\BAL_MASI_TFL").Value
        liquid_density_lb_gal = mass_flow_kg_s / pump_flowrate[i-1] * kg_m3_to_lb_gal

        # - Head ≤ 3200 ft: centrifugal formula (function of gpm)
        # - Head > 3200 ft: reciprocating, fixed efficiency 0.9
        if pump_head[i-1] <= 3200:
            gpm = pump_flowrate[i-1] * m3_s_to_gpm
            if gpm < 50:
                eff = (-0.316 + 0.24015*np.log(50) - 0.01199*(np.log(50))**2) * (gpm/50)**0.6
            else:
                eff = (-0.316 + 0.24015*np.log(gpm) - 0.01199*(np.log(gpm))**2)
        else:
            eff = 0.9

        # Break horsepower [HP]
        pump_break_horsepower[i-1] = pump_head[i-1] * pump_flowrate[i-1] * m3_s_to_gpm * liquid_density_lb_gal / (33000 * eff)

    return pump_head, pump_flowrate, pump_size_factor, pump_break_horsepower, pump_electricity_W





def vessel_sizing(liquid_flowrate, length_to_diameter_ratio, hold_up_time, liquid_fill=0.8):
    """
    Calculate the diameter, length, and volume of a vessel.

    Args:
    liquid_flowrate (float): Flow rate in m³/s
    length_to_diameter_ratio (float): Ratio of length to diameter
    hold_up_time (float): Desired hold-up time in s
    liquid_fill (float): Fraction of the vessel filled with liquid (default is 0.8)

    Returns:
    Diameter (m), Length (m), Volume (m³)
    """
    


    # Calculate diameter
    D_cubed = (4 * hold_up_time * liquid_flowrate) / (math.pi * liquid_fill * length_to_diameter_ratio)
    D = (D_cubed)**(1/3)

    # Calculate length using the length to diameter ratio
    L = D * length_to_diameter_ratio

    # Calculate the volume of the vessel
    Volume = (math.pi * (D**2) * L * liquid_fill) / 4

    return D, L, Volume

