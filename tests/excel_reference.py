"""Results of the Excel STREAM model, used to check that pySTREAM gives the same numbers.

The values were copied from the results Excel had last calculated in
'STREAM_bioPlus_2015baseyear for validation.xlsm' ('Base' and 'Scenario' columns).
This workbook is not included in pySTREAM; STREAM is described at
https://www.esymodels.man.dtu.dk/stream
The sheet and cell of each value are given in the comments.

The Excel scenario uses a CO2 price of 32.7 EUR/t ('Basic Information'!E58), while
pySTREAM's scenario data uses another price, so scenario CO2 costs are compared per
tonne of CO2 (cost divided by price).
"""

# 'Basic Information'!C58 and E58 [EUR/t CO2]
EXCEL_CO2_PRICE = {'BY': 5.0, 'SC': 32.7}

# Total fuel costs in the base year [mill. EUR]: sum of the fuel cost blocks in 'Economic balance' row 9
# (H electricity, T district heat, AF residential, AR industry, BD tertiary, BP agriculture,
# BY transport, CK fuel refinery input, CN energy branch). Excel's scenario uses other fuel prices.
EXCEL_FUEL_COSTS_TOTAL_BY = 3916.8391357024757

EXCEL = {
    'BY': {
        # 'Electricity Demand'!B23 and 'DH Demand'!B22 [PJ]
        'total_electricity': 130.10253307107897,
        'total_DH_demand': 118.96155294117648,
        # 'Energy balance'!AU13:AU26, 'Fuel consumption Total (incl fuel refinery)' [PJ]
        'primary_fuel': {
            'Bioenergy': 138.06425732368638,
            'Coal': 93.63851652173912,
            'District Heating': -50.41145229224877,
            'Electricity': 28.332533071078977,
            'Geothermal': 0.05982905982905984,
            'Hydro': 0.06,
            'Natural Gas': 97.60145216690766,
            'Nuclear': 0.0,
            'Oil': 218.1286776314081,
            'Solar': 3.5877299999999996,
            'Waste': 53.31627159850655,
            'Wave': 0.0,
            'Wind': 49.6,
        },
        # 'Environmental balance' rows 13-26 [Mt CO2]
        'CO2': {
            'Electricity production no CCS': {  # column D
                'Bioenergy': -0.17814285714285716,
                'Coal': 5.280760869565217,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.7780425531914895,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.4917272727272727,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Electricity production w. CCS': {  # column J
                'Bioenergy': -0.17814285714285716,
                'Coal': 5.280760869565217,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.7780425531914895,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.4917272727272727,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'District Heat': {  # column AE
                'Bioenergy': -0.08766990291262135,
                'Coal': 3.0898749999999997,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 1.2814757281553397,
                'Nuclear': 0.0,
                'Oil': 0.12949999999999998,
                'Solar': 0.0,
                'Waste': 0.46796561604584525,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Fuel Refinery': {  # column AB
                'Bioenergy': -0.9269368042960275,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.0,
                'Nuclear': 0.0,
                'Oil': 4.0938022408927,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Residential': {  # column S
                'Bioenergy': 0.0,
                'Coal': 0.03940470588235294,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 1.692019512195122,
                'Nuclear': 0.0,
                'Oil': 0.9911375,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Tertiary': {  # column M
                'Bioenergy': 0.0,
                'Coal': 0.004678470588235295,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.5003736585365853,
                'Nuclear': 0.0,
                'Oil': 0.2439854,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Industry': {  # column P
                'Bioenergy': 0.0,
                'Coal': 0.5155966666666666,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 1.8613894634146344,
                'Nuclear': 0.0,
                'Oil': 1.620605025,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Agriculture': {  # column V
                'Bioenergy': 0.0,
                'Coal': 0.08422811764705883,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.1000747317073171,
                'Nuclear': 0.0,
                'Oil': 1.293511675,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Transport': {  # column Y
                'Bioenergy': 0.0,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.0,
                'Nuclear': 0.0,
                'Oil': 8.787870311462907,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
        },
        # 'Economic balance' row 9 (totals) [mill. EUR]
        'costs': {
            'capital': {
                'Electricity': 881.2702604408039,  # B9
                'District Heat': 372.4801710346117,  # N9
                'Residential': 329.70773122624104,  # Z9
                'Industry': 55.27126766289767,  # AL9
                'Tertiary': 39.850530876417245,  # AX9
                'Agriculture': 28.387658667101675,  # BJ9
                'Fuel Refinery': 734.9572045982017,  # CE9
                'Transport': 2455.4180669448506,  # BV9 (vehicles: investment and maintenance)
            },
            'O&M': {
                'Electricity': 460.558683791185,  # E9
                'District Heat': 323.47261995066947,  # Q9
                'Residential': 180.70358922550153,  # AC9
                'Industry': 19.945646984126984,  # AO9
                'Tertiary': 45.88253003823296,  # BA9
                'Agriculture': 11.058620119863013,  # BM9
            },
            'CO2': {
                'District Heat': 24.84408172100592,  # W9
                'Residential': 13.612808590387374,  # AI9
                'Industry': 19.987955775406505,  # AU9
                'Tertiary': 3.745187645624103,  # BG9
                'Agriculture': 7.389072621771879,  # BS9
                'Transport': 43.93935155731454,  # CB9
            },
        },
    },
    'SC': {
        # 'Electricity Demand'!D23 and 'DH Demand'!D22 [PJ]
        'total_electricity': 124.89312255796283,
        'total_DH_demand': 103.32921060197256,
        # 'Energy balance'!AW13:AW26, 'Fuel consumption Total (incl fuel refinery)' [PJ]
        'primary_fuel': {
            'Bioenergy': 283.6041439560635,
            'Coal': 99.23931958998737,
            'District Heating': -2.4868995751603507e-14,
            'Electricity': 1.2702516845385808e-11,
            'Geothermal': 1.0628814449774613,
            'Hydro': 0.0,
            'Natural Gas': 27.089007829205443,
            'Nuclear': 0.0,
            'Oil': 20.043173484989246,
            'Solar': 9.893739220778956,
            'Waste': 37.441193058057856,
            'Wave': 0.0,
            'Wind': 59.69891258270017,
        },
        # 'Environmental balance' rows 13-26 [Mt CO2]
        'CO2': {
            'Electricity production no CCS': {  # column F
                'Bioenergy': -0.7160539026655808,
                'Coal': 9.317438033038396,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 1.3641822109508466,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.6739414750450413,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Electricity production w. CCS': {  # column L
                'Bioenergy': -0.7160539026655808,
                'Coal': 0.9317438033038385,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.1364182210950846,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.6739414750450413,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'District Heat': {  # column AG
                'Bioenergy': 0.0,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.0,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Fuel Refinery': {  # column AD
                'Bioenergy': -3.5269717699224503,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.0,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Residential': {  # column U
                'Bioenergy': 0.0,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.0,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Tertiary': {  # column O
                'Bioenergy': 0.0,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.0,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Industry': {  # column R
                'Bioenergy': 0.0,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.04122086324250116,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Agriculture': {  # column X
                'Bioenergy': 0.0,
                'Coal': 0.11610245053726764,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.11570345056640718,
                'Nuclear': 0.0,
                'Oil': 1.5449946228012543,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
            'Transport': {  # column AA
                'Bioenergy': -0.3503975584167737,
                'Coal': 0.0,
                'District Heating': 0.0,
                'Electricity': 0.0,
                'Geothermal': 0.0,
                'Hydro': 0.0,
                'Natural Gas': 0.0,
                'Nuclear': 0.0,
                'Oil': 0.0,
                'Solar': 0.0,
                'Waste': 0.0,
                'Wave': 0.0,
                'Wind': 0.0,
            },
        },
        # 'Economic balance' row 9 (totals) [mill. EUR]
        'costs': {
            'capital': {
                'Electricity': 902.0514973374434,  # D9
                'District Heat': 38.105904632661776,  # P9
                'Residential': 278.8756624935956,  # AB9
                'Industry': 89.34056373829856,  # AN9
                'Tertiary': 51.52139075358225,  # AZ9
                'Agriculture': 38.51679405007549,  # BL9
                'Fuel Refinery': 722.7022814926677,  # CG9
                'Transport': 5797.405070876311,  # BX9 (vehicles: investment and maintenance)
            },
            'O&M': {
                'Electricity': 569.8097209473024,  # G9
                'District Heat': 16.367864217181737,  # S9
                'Residential': 106.3580960501556,  # AE9
                'Industry': 31.818841208833568,  # AQ9
                'Tertiary': 24.677479459003894,  # BC9
                'Agriculture': 17.453776737727054,  # BO9
            },
            'CO2': {
                'District Heat': 0.0,  # Y9
                'Residential': 0.0,  # AK9
                'Industry': 1.347922228029788,  # AW9
                'Tertiary': 0.0,  # BI9
                'Agriculture': 58.10137713169118,  # BU9
                'Transport': 0.0,  # CD9
            },
        },
    },
}
