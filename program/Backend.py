
import os
import pandas as pd
import numpy as np
import numpy_financial as npf

#The pySTREAM folder (this file is in its 'program' folder), so the data is found no matter which folder
#pySTREAM is started from. The data, figures and error log are in the pySTREAM folder, where students look for them.
PROGRAM_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(PROGRAM_DIR, 'data')


#Global constants
sectors = ['Residential', 'Tertiary', 'Industry', 'Agriculture']
tech_types = ['Electricity appliances', 'Electric Heater', 'District Heating', 'Coal Boiler', 'Oil Boiler', 'Gas Boiler', 'Wood Pellet Boiler', 'Solar heating', 'Heatpumps']
transport_types = ['Car', 'Bus', 'Train', 'Aviation & ferries', 'Bike', 'Trucks and cargo vans', 'Train Goods', 'Shipping', 'Air transport', 'Fishery', 'Agriculture']
fuels = ['Electricity',	'Gasoline',	'Diesel', 'Natural Gas', 'Methanol', 'Ethanol', 'Biodiesel', 'Hydrogen', 'UPGR Biogas']
FUEL_TYPES = {'Bioenergy': {'Biomass (Energy crops)', 'Biomass (Manure)', 'Biomass (Straw,woodwaste)', 'Biogas', 'Ethanol', 'Methanol', 'Biodiesel', 'UPGR Biogas'},
              'Coal': {'Coal', 'Lignite', 'Cheap coal'},
              'Oil': {'Oil', 'Crude Oil', 'Gasoline', 'Diesel'},
              'District Heating': {'District Heating', 'Imported heat'}, 
              'Electricity': {'Electricity', 'Imported electricity'},
              'Geothermal': {'Geothermal'},
              'Hydro': {'Hydropower'},
              'Waste': {'Municipal Waste', 'Municipal Waste, RE', 'Municipal Waste, non-RE'}, 
              'Natural Gas': {'Natural Gas'},
              'Nuclear': {'Nuclear', 'Nuclear - Uranium', 'Thorium'},
              'Solar': {'Solar Heating', 'Solar Power', 'Surplus Heat'}, 
              'Wave': {'Wave'}, 
              'Wind': {'Wind'},
               }
FUEL_TYPES_PRIMARY = {'Bioenergy': {'Biomass (Energy crops)', 'Biomass (Manure)', 'Biomass (Straw,woodwaste)'},
                      'Coal': {'Coal', 'Lignite', 'Cheap coal'},
                      'Oil': {'Oil', 'Crude Oil'},
                      'District Heating': {'District Heating', 'Imported heat'}, 
                      'Electricity': {'Electricity', 'Imported electricity'},
                      'Geothermal': {'Geothermal'},
                      'Hydro': {'Hydropower'},
                      'Waste': {'Municipal Waste', 'Municipal Waste, RE', 'Municipal Waste, non-RE'}, 
                      'Natural Gas': {'Natural Gas'},
                      'Nuclear': {'Nuclear', 'Nuclear - Uranium', 'Thorium'},
                      'Solar': {'Solar Heating', 'Solar Power', 'Surplus Heat'}, 
                      'Wave': {'Wave'}, 
                      'Wind': {'Wind'},
                      }

CATEGORIES = {'Energy Carrier': {'Electricity', 'District Heating'},
              'Fossil': {'Coal', 'Natural Gas', 'Oil'},
              'Renewable': {'Bioenergy', 'Geothermal', 'Hydro', 'Solar', 'Waste', 'Wave', 'Wind'},
              'Nuclear': {'Nuclear'}
              }
#Model assumptions (grid losses, discount rate, CO2 capture rate, ...) are read from data/Assumptions.xlsx
ASSUMPTIONS = pd.read_excel(os.path.join(DATA_DIR, 'Assumptions.xlsx'), sheet_name='Assumptions').set_index('Assumption')['Value']
GL_el_COEF = ASSUMPTIONS['Grid losses, electricity']
GL_DH_COEF = ASSUMPTIONS['Grid losses, district heating']



#Load Base Year data
#==============================================================================
BYData = os.path.join(DATA_DIR, 'Base Year Data.xlsx')
df_BY = pd.read_excel(BYData, sheet_name='Base Year')
Base_Year = df_BY.iat[0,0]
  
df_FEC_BY = pd.read_excel(BYData, sheet_name='Final Energy Consumption')
df_FEC_BY.set_index('Tech Type', inplace=True)
df_FEC_BY_total = pd.DataFrame(df_FEC_BY.sum(axis=0, numeric_only=True), columns=["Total"])
df_FEC_BY['Total by type'] = df_FEC_BY.sum(axis=1, numeric_only=True)

df_Power_BY = pd.read_excel(BYData, sheet_name='Power Production')
df_Power_BY['Tech Types'] = df_Power_BY['Tech Type'] #Adding row of names for tech types to be used in tables
df_Power_BY.set_index('Tech Type', inplace=True) #Using Tech Type as the index of the rows, so rows can be refered to by tech type
BY_total_power = df_Power_BY['Amount'].sum()

df_Heat_BY = pd.read_excel(BYData, sheet_name='Heat Production')
df_Heat_BY['Tech Types'] = df_Heat_BY['Tech Type'] #Adding row of names for tech types to be used in tables
df_Heat_BY.set_index('Tech Type', inplace=True) #Using Tech Type as the index of the rows, so rows can be refered to by tech type
BY_total_heat = df_Heat_BY['Amount'].sum()

df_Transport_BY = pd.read_excel(BYData, sheet_name='Transport')
df_Transport_BY.set_index("Type", inplace=True)
df_TransportFuel_BY = pd.read_excel(BYData, sheet_name='Fuels Transportation')
df_TransportFuel_BY.set_index("Type", inplace=True)
df_TransportWork_BY = pd.read_excel(BYData, sheet_name='Transport Work')
df_TransportWork_BY.set_index("Transport Type", inplace=True)

df_SpecificEnergyDemand = pd.read_excel(BYData, sheet_name='Specific Energy Demand')
df_SpecificEnergyDemand.set_index("Transport type", inplace=True)

df_Resources_BY = pd.read_excel(BYData, sheet_name='Resources Available')

df_Economic = pd.read_excel(BYData, sheet_name='Economic Growth')
sectors_extended = df_Economic['Sector'].to_numpy()

df_TechEff_BY = pd.read_excel(BYData, sheet_name='Technology Efficiencies')
df_TechEff_BY.set_index("Fuel", inplace=True)

df_FuelRefinery_BY = pd.read_excel(BYData, sheet_name='Fuel Refinery Fuel Input Factor')
df_FuelRefinery_BY.set_index("Fuel Type", inplace=True)

df_FuelPrice_BY = pd.read_excel(BYData, sheet_name='Fuel Price')
df_FuelPrice_BY.set_index("Fuel Type", inplace=True)

df_FuelRefineryCost_BY = pd.read_excel(BYData, sheet_name='Fuel Refinery Cost')
df_FuelRefineryCost_BY.set_index("Fuel Type", inplace=True)

df_CO2 = pd.read_excel(BYData, sheet_name='CO2')
df_CO2.set_index('Fuel Type', inplace=True)

df_SectorsDetails_BY = pd.read_excel(BYData, sheet_name='Sectors Details')
df_SectorsDetails_BY.set_index("Tech Type", inplace=True)

#Costs of vehicles (investment and maintenance) per person km or tonne km, per transport type and fuel
df_TransportCosts_BY = pd.read_excel(BYData, sheet_name='Transport Costs', index_col=0)
#==============================================================================

SCData = os.path.join(DATA_DIR, 'Scenario Data.xlsx')
DEFAULT_SCData = os.path.join(DATA_DIR, 'Default Scenario Data.xlsx') #the scenario that comes with pySTREAM, never changed

df_Resources_SC = pd.read_excel(SCData, sheet_name='Resources Available')


#Base year calculations
def BY_calculations():
    #calculations that are based on data from the base year
    #will update if scenario year is changed
    
    df_SC = pd.read_excel(SCData, sheet_name='Scenario Year')
    Scenario_Year = df_SC.iat[0,0]
    
    #Calculations for energy demand for Scenario year in the sectors: Residential, Tertiary, Industry, and Agriculture
    #==========================================================================
    def calculate_egr():
        #Calculate economic growth rate (egr) for all sectors
        #----------------------------------------------------------------------------------
        egr = np.full(shape=len(sectors_extended),fill_value=1.0)
        
        for i in range(len(sectors_extended)):
            
            for j in range(Scenario_Year - Base_Year):
                egr[i] = egr[i] * (1 + df_Economic['Economic growth'][i])
                
        return egr
        #----------------------------------------------------------------------------------
        
    def calculate_engr():
        #Calculate energy growth rate (engr) for all sectors
        #----------------------------------------------------------------------------------
        engr = np.full(shape=len(sectors_extended),fill_value=1.0)
        for i in range(len(sectors_extended)):
            for j in range(Scenario_Year - Base_Year):
                engr[i] = engr[i] * (1 + df_Economic['Economic growth'][i] * df_Economic['Structure/intensity factor'][i])
                
        return engr
        #----------------------------------------------------------------------------------
    
    def calculate_NEC():
        #Calculate Net Energy Consumption (NEC) for the base year for all sectors
        #----------------------------------------------------------------------------------
        technologies = df_TechEff_BY.index.to_numpy()
        BY_NEC = np.empty(shape=(len(technologies),len(sectors)))
        for j in range(len(sectors)):
            for m in range(len(technologies)):
                BY_NEC[m,j] = df_TechEff_BY.iloc[m][sectors[j]]*df_FEC_BY.iloc[m][sectors[j]]
        
        #Total Net Energy Consumption for each sector for electricity and heat
        BY_NEC_t_el = BY_NEC[0]
        BY_NEC_t_heat = sum(BY_NEC[1:])
        
        #beta = share of each technology in NEC for the different energy carriers
        beta_el = BY_NEC[0]/BY_NEC_t_el
        beta_heat = np.empty(shape=(len(technologies)-1,len(sectors)))
        for j in range(len(sectors)):
            beta_heat[:,j] = BY_NEC[1:,j]/BY_NEC_t_heat[j]
            
        df_beta_BY = pd.DataFrame(np.concatenate(([beta_el], beta_heat)), columns=sectors, index=tech_types)
    
        return BY_NEC_t_el, BY_NEC_t_heat, df_beta_BY
        #----------------------------------------------------------------------------------
        
    def calculate_NEC_tFE(engr, BY_NEC_t_el, BY_NEC_t_heat):
        #Calculate Total Net Energy Consumption (NEC_tFE) for Scenario year considering frozen efficiencies (FE)
        #----------------------------------------------------------------------------------
        #create array with only electricity engr
        engr_el = np.array([engr[0],engr[2],engr[4],engr[6]])
    
        SC_NEC_tFE_el = engr_el * BY_NEC_t_el
        
        #create array with only heat engr
        engr_heat = np.array([engr[1],engr[3],engr[5],engr[7]])
        
        SC_NEC_tFE_heat = engr_heat * BY_NEC_t_heat
        
        return SC_NEC_tFE_el, SC_NEC_tFE_heat
        #----------------------------------------------------------------------------------
    
    egr = calculate_egr()
    engr = calculate_engr()
    BY_NEC_t_el, BY_NEC_t_heat, df_beta_BY = calculate_NEC()
    SC_NEC_tFE_el, SC_NEC_tFE_heat = calculate_NEC_tFE(engr, BY_NEC_t_el, BY_NEC_t_heat)
    #==================================================================================
    return SC_NEC_tFE_el, SC_NEC_tFE_heat, df_beta_BY, engr

#Base year calculations  
SC_NEC_tFE_el, SC_NEC_tFE_heat, df_beta_BY, engr = BY_calculations()

#Transport demand in the scenario year: base year demand times the energy growth rate (engr)
#Used both when loading the data and when the scenario year is changed in the GUI
def set_SC_transport_demand(df_Transport_SC, engr):
    #Not rounded to whole numbers, as in the Excel model (the column is read from Excel as whole numbers, so make it decimal first)
    df_Transport_SC['Amount'] = df_Transport_SC['Amount'].astype(float)
    df_Transport_SC.loc['Passenger transport','Amount'] = df_Transport_BY.iloc[0]['Amount']*engr[-4]
    df_Transport_SC.loc['Freight transport','Amount'] = df_Transport_BY.iloc[1]['Amount']*engr[-3]
    df_Transport_SC.loc['Fishery','Amount'] = df_Transport_BY.iloc[2]['Amount']*engr[-2]
    df_Transport_SC.loc['Agriculture','Amount'] = df_Transport_BY.iloc[3]['Amount']*engr[-1]

#Load Scenario year data and do calculations for scenario year
#======================================================================================
def load_SCdata():
    #Load data from excel file
    #----------------------------------------------------------------------------------
    df_SC = pd.read_excel(SCData, sheet_name='Scenario Year')
    Scenario_Year = df_SC.iat[0,0]
    
    df_beta_SC = pd.read_excel(SCData, sheet_name='Final Energy Consumption')
    df_beta_SC.set_index('Tech Type', inplace=True)
    
    df_Power_SC = pd.read_excel(SCData, sheet_name='Electricity Production')
    df_Power_SC['Tech Types'] = df_Power_SC['Tech Type'] #Adding row of names for tech types to be used in tables
    df_Power_SC.set_index('Tech Type', inplace=True) #Using Tech Type as the index of the rows, so rows can be refered to by tech type
    
    df_Heat_SC = pd.read_excel(SCData, sheet_name='District Heat Boiler Production')
    df_Heat_SC['Tech Types'] = df_Heat_SC['Tech Type'] #Adding row of names for tech types to be used in tables
    df_Heat_SC.set_index('Tech Type', inplace=True) #Using Tech Type as the index of the rows, so rows can be refered to by tech type
    
    df_Transport_SC = pd.read_excel(SCData, sheet_name='Transport')
    df_Transport_SC.set_index("Type", inplace=True)
    set_SC_transport_demand(df_Transport_SC, engr)
    df_TransportFuel_SC = pd.read_excel(SCData, sheet_name='Fuels Transportation')
    df_TransportFuel_SC.set_index("Type", inplace=True)
    df_TransportWork_SC = pd.read_excel(SCData, sheet_name='Transport Work')
    df_TransportWork_SC.set_index("Transport Type", inplace=True)
    df_TransportImprovement_SC = pd.read_excel(SCData, sheet_name='Technology Improvements')
    df_TransportImprovement_SC.set_index("Type", inplace=True)
    
    
    df_EnergySavings_SC = pd.read_excel(SCData, sheet_name='Energy Savings')
    
    df_TechEff_SC = pd.read_excel(SCData, sheet_name='Technology Efficiencies')
    df_TechEff_SC.set_index("Fuel", inplace=True)
    
    df_FuelRefinery_SC = pd.read_excel(SCData, sheet_name='Fuel Refinery Fuel Input Factor')
    df_FuelRefinery_SC.set_index("Fuel Type", inplace=True)
    
    df_FuelPrice_SC = pd.read_excel(SCData, sheet_name='Fuel Price')
    df_FuelPrice_SC.set_index("Fuel Type", inplace=True)

    df_FuelRefineryCost_SC = pd.read_excel(SCData, sheet_name='Fuel Refinery Cost')
    df_FuelRefineryCost_SC.set_index("Fuel Type", inplace=True)
    
    df_SectorsDetails_SC = pd.read_excel(SCData, sheet_name='Sectors Details')
    df_SectorsDetails_SC.set_index("Tech Type", inplace=True)

    df_CO2_SC = pd.read_excel(SCData, sheet_name='CO2 Price')
    co2_price_SC = df_CO2_SC.iat[0,0]

    #Costs of vehicles per person km or tonne km. Scenario files saved before this sheet was added use the default scenario's.
    with pd.ExcelFile(SCData) as workbook:
        transport_costs_file = SCData if 'Transport Costs' in workbook.sheet_names else DEFAULT_SCData
    df_TransportCosts_SC = pd.read_excel(transport_costs_file, sheet_name='Transport Costs', index_col=0)
  

    return Scenario_Year, df_beta_SC, df_Power_SC, df_Heat_SC, df_Transport_SC, df_TransportFuel_SC, df_TransportWork_SC, df_TransportImprovement_SC, df_EnergySavings_SC, df_TechEff_SC, df_FuelRefinery_SC, df_FuelPrice_SC, df_FuelRefineryCost_SC, df_SectorsDetails_SC, co2_price_SC, df_TransportCosts_SC
    
    #----------------------------------------------------------------------------------
Scenario_Year, df_beta_SC, df_Power_SC, df_Heat_SC, df_Transport_SC, df_TransportFuel_SC, df_TransportWork_SC, df_TransportImprovement_SC, df_EnergySavings_SC, df_TechEff_SC, df_FuelRefinery_SC, df_FuelPrice_SC, df_FuelRefineryCost_SC, df_SectorsDetails_SC, co2_price_SC, df_TransportCosts_SC = load_SCdata()


def SC_calculations(SC_NEC_tFE_el, SC_NEC_tFE_heat):
    SC_NEC_tES_el, SC_NEC_tES_heat = calculate_NEC_tES(SC_NEC_tFE_el, SC_NEC_tFE_heat)
    SC_NEC_FE_el, SC_NEC_FE_heat, SC_NEC_ES_el, SC_NEC_ES_heat = calculate_SC_NEC(SC_NEC_tFE_el, SC_NEC_tFE_heat, SC_NEC_tES_el, SC_NEC_tES_heat)
    SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES = calculate_SC_FEC(SC_NEC_FE_el, SC_NEC_FE_heat, SC_NEC_ES_el, SC_NEC_ES_heat)
    
    return SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES
    
#Calculate Total Net Energy Consumption (NEC_tES) for Scenario year considering energy savings (ES)
#----------------------------------------------------------------------------------
def calculate_NEC_tES(SC_NEC_tFE_el, SC_NEC_tFE_heat):
    
    ES = df_EnergySavings_SC['Final energy savings compared to frozen efficiency'].to_numpy()
    ES_el = ES[0:4]
    SC_NEC_tES_el = SC_NEC_tFE_el * (1 - ES_el)
    ES_heat = ES[4:]
    SC_NEC_tES_heat = SC_NEC_tFE_heat * (1 - ES_heat)
    
    return SC_NEC_tES_el, SC_NEC_tES_heat
#----------------------------------------------------------------------------------
    
#Calculate Net Energy Consumption (NEC) for Scenario year, frozen efficiencies (FE) and energy savings (ES)
#----------------------------------------------------------------------------------
def calculate_SC_NEC(SC_NEC_tFE_el, SC_NEC_tFE_heat, SC_NEC_tES_el, SC_NEC_tES_heat):
    #beta for scenario is based on user input
    SC_beta_el = df_beta_SC.iloc[0][:].to_numpy()
    SC_beta_heat = df_beta_SC.iloc[1:,:].to_numpy() 
    
    SC_NEC_FE_el = SC_beta_el * SC_NEC_tFE_el
    SC_NEC_FE_heat = SC_beta_heat * SC_NEC_tFE_heat
   
    SC_NEC_ES_el = SC_beta_el * SC_NEC_tES_el
    SC_NEC_ES_heat = SC_beta_heat * SC_NEC_tES_heat
    
    return SC_NEC_FE_el, SC_NEC_FE_heat, SC_NEC_ES_el, SC_NEC_ES_heat
#----------------------------------------------------------------------------------

#Calculate Final Energy Consumption (FEC) and total Final Energy Consumption (FEC_t) for Scenario year (SC), frozen efficiencies (FE) and energy savings (ES)
#----------------------------------------------------------------------------------
def calculate_SC_FEC(SC_NEC_FE_el, SC_NEC_FE_heat, SC_NEC_ES_el, SC_NEC_ES_heat):
    
    SC_TechEff_el = df_TechEff_SC.iloc[0][1:].to_numpy()
    SC_TechEff_heat = df_TechEff_SC.iloc[1:,1:].to_numpy() 
      
    SC_FEC_FE_el = SC_NEC_FE_el / SC_TechEff_el
    SC_FEC_tFE_el = SC_FEC_FE_el 
    
    SC_FEC_FE_heat = SC_NEC_FE_heat / SC_TechEff_heat
    SC_FEC_tFE_heat = sum(SC_FEC_FE_heat)
    
    SC_FEC_ES_el = SC_NEC_ES_el / SC_TechEff_el
    SC_FEC_tES_el = SC_FEC_ES_el 

    SC_FEC_ES_heat = SC_NEC_ES_heat / SC_TechEff_heat
    SC_FEC_tES_heat = sum(SC_FEC_ES_heat)
    SC_FEC_ES = np.vstack((SC_FEC_ES_el,SC_FEC_ES_heat))
    SC_FEC_tFE = SC_FEC_tFE_el + SC_FEC_tFE_heat
    
    #Total final energy consumption for each sector in the scenario year
    SC_FEC_tES = SC_FEC_tES_el + SC_FEC_tES_heat
    
    return SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES
#----------------------------------------------------------------------------------
        
SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES = SC_calculations(SC_NEC_tFE_el, SC_NEC_tFE_heat)

#Dataframe containing FEC for base year and scenario
#----------------------------------------------------------------------------------
def FEC_dataframe(SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES):
    #Creating a dataframe for the Final Energy Consumption (FEC) for the scenario year
    df_FEC_SC = pd.DataFrame(SC_FEC_ES, columns=sectors, index=tech_types)
    df_FEC_SC.index.rename('Tech Type', inplace=True)
    df_FEC_SC['Total by type'] = df_FEC_SC.sum(axis=1)
    
    #Create dataframe for total consumption per sector for scenario year
    df_FEC_SC_total = pd.DataFrame(SC_FEC_tES, index=sectors, columns=["Total"])
    
    return df_FEC_SC.astype(float), df_FEC_SC_total

df_FEC_SC, df_FEC_SC_total = FEC_dataframe(SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES)
#----------------------------------------------------------------------------------

#Calculate Transport demandand fuel consumption
#----------------------------------------------------------------------------------
def calculate_transport(year, df_FEC, df_Transport, df_TransportWork, df_TransportFuel):
    #Transport energy consumption [TJ]
    df_transport_consumption = pd.DataFrame(index=transport_types, columns=['Energy consumption'])
    
    #Persons transport (Car, Bus, Train, Aviation and ferries, Bike [0:4])
    for i in df_transport_consumption[:5].index:
        #Calculated for each transport type as Passenger transport [mio.person km] * transport work share * (sum of (Fuel share * Specific energy demand))      
        utilization = (df_TransportWork_BY.loc[i,'Utilisation degree']/df_TransportWork.loc[i,'Utilisation degree'])
        improvement_factor = 1
        if year == 'Scenario Year':
            improvement_factor = (1 - df_TransportImprovement_SC.loc[i])
        df_transport_consumption.loc[i,'Energy consumption'] = utilization * df_Transport.loc['Passenger transport','Amount']*df_TransportWork.loc[i,'Allocation of transport work'] * ((df_TransportFuel.loc[i] * (df_SpecificEnergyDemand.loc[i] * improvement_factor)).sum())

    #Goods transport (Trucks and cargo vans, Train (Goods), Shipping, Air transport [5:8])
    for i in df_transport_consumption[5:9].index:
        #Calculated for each transport type as Freight transport [mio.tonnes km] * transport work share * (sum of (Fuel share * Specific energy demand))
        utilization = (df_TransportWork_BY.loc[i,'Utilisation degree']/df_TransportWork.loc[i,'Utilisation degree'])
        improvement_factor = 1
        if year == 'Scenario Year':
            improvement_factor = (1 - df_TransportImprovement_SC.loc[i])
        df_transport_consumption.loc[i,'Energy consumption'] = utilization * df_Transport.loc['Freight transport','Amount']*df_TransportWork.loc[i,'Allocation of transport work'] * ((df_TransportFuel.loc[i] * (df_SpecificEnergyDemand.loc[i] * improvement_factor)).sum())          

    #Fishing and agriculture transport (Fishing fleet, Agricultural machinery)
    df_transport_consumption.loc['Fishery', 'Energy consumption'] = df_Transport.loc['Fishery','Amount']
    df_transport_consumption.loc['Agriculture', 'Energy consumption'] = df_Transport.loc['Agriculture','Amount']

    #Calculate Total transport consumtion [TJ]
    transport_consumption_total = df_transport_consumption['Energy consumption'].sum()
    df_transport_consumption_fuels = pd.DataFrame(index=fuels, columns=['Share of total transport consumption'])
    
    #Total fuel shares for each fuel type
    for i in range(df_TransportFuel.shape[1]):
        df_transport_consumption_fuels.loc[fuels[i], 'Share of total transport consumption'] = (df_TransportFuel[fuels[i]] * df_transport_consumption['Energy consumption']).sum() / transport_consumption_total
    
    
    #Fuel consumption for each transport type and fuel type
    utilization = pd.Series(index=df_transport_consumption[:9].index)
    for i in df_transport_consumption[:9].index:
        utilization[i] = df_TransportWork_BY.loc[i,'Utilisation degree']/df_TransportWork.loc[i,'Utilisation degree']
    
    for i in range(df_TransportFuel.shape[1]):
        
        if year == 'Scenario Year': #in scenario year the transport fuel consumption is calculated including utilization degree
            improvement_factor = (1 - df_TransportImprovement_SC[fuels[i]])
            
            #person
            share_person = 0
            for transportType in df_transport_consumption[:4].index:
               
                share_person += df_TransportFuel.loc[transportType,fuels[i]] * (improvement_factor.loc[transportType] * df_SpecificEnergyDemand.loc[transportType,fuels[i]]) * df_TransportWork.loc[transportType,'Allocation of transport work'] * utilization[transportType]
            share_person = share_person * df_Transport.loc['Passenger transport','Amount']
            
            #goods
            share_goods = 0
            for transportType in df_transport_consumption[5:9].index:
                share_goods += df_TransportFuel.loc[transportType,fuels[i]]* (improvement_factor.loc[transportType] * df_SpecificEnergyDemand.loc[transportType,fuels[i]]) * df_TransportWork.loc[transportType,'Allocation of transport work'] * utilization[transportType]
            share_goods = share_goods * df_Transport.loc['Freight transport','Amount']
            
            #fishery and agriculture
            share_agriculture = df_Transport.loc['Fishery','Amount'] * df_TransportFuel.loc['Fishery',fuels[i]]
            share_fishery = df_Transport.loc['Agriculture','Amount'] * df_TransportFuel.loc['Agriculture',fuels[i]]
         
            df_transport_consumption_fuels.loc[fuels[i],'Total transport consumption'] = share_person + share_goods + share_agriculture + share_fishery
            
        else: #Base year
            df_transport_consumption_fuels.loc[fuels[i],'Total transport consumption'] = transport_consumption_total * df_transport_consumption_fuels.loc[fuels[i],'Share of total transport consumption']

    return df_transport_consumption, df_transport_consumption_fuels
#----------------------------------------------------------------------------------

df_transport_consumption_BY, df_transport_consumption_fuels_BY = calculate_transport('Base Year', df_FEC_BY, df_Transport_BY, df_TransportWork_BY, df_TransportFuel_BY)
df_transport_consumption_SC, df_transport_consumption_fuels_SC = calculate_transport('Scenario Year', df_FEC_SC, df_Transport_SC, df_TransportWork_SC, df_TransportFuel_SC)    

#Calculate Combined heat and power (CHP) production
#----------------------------------------------------------------------------------
def calculate_CHP_heat(df_Power, year):
    productionUnits = df_Power.index.values
    total_electricity_production =  df_Power['Amount'] 
    
    share_condensing_power = []
    for i in productionUnits:
        if df_Power['Mode'][i] == 'Backpressure':
            x = 0
        elif df_Power['Mode'][i] == 'Fuel free' or df_Power['Mode'][i] == 'Condensing':
            x = 1
        else: 
            if year == "Base year":
                x = ASSUMPTIONS['Condensing share of extraction plants, base year']
            if year == "Scenario year":
                x = ASSUMPTIONS['Condensing share of extraction plants, scenario year']
        
        share_condensing_power.append(x)
        
    #Electricity production in condensing or non CHP mode
    electricity_production_conORnonCHP = total_electricity_production * share_condensing_power
    
    electricity_production = total_electricity_production - electricity_production_conORnonCHP
    
    fuel_consumption_CHP = electricity_production / df_Power['Electric efficiency in CHP mode']
    
    df_Power_chp = pd.DataFrame(index=productionUnits, columns=['CHP_heat_production'])
    for i in productionUnits:
        if df_Power['Mode'][i] != 'Condensing' and df_Power['Mode'][i] != 'Fuel free':
            df_Power_chp.loc[i, 'CHP_heat_production'] = (df_Power['Overall Efficiency (CHP: el+heat)'][i] - df_Power['Electric efficiency in CHP mode'][i]) * fuel_consumption_CHP[i]
        else:
            df_Power_chp.loc[i, 'CHP_heat_production'] = None
            
    return df_Power_chp, electricity_production_conORnonCHP, fuel_consumption_CHP
#----------------------------------------------------------------------------------


df_Power_BY['CHP heat production'], df_Power_BY['Electricity production (condensing or nonCHP)'], df_Power_BY['Fuel consumption in CHP-mode'] = calculate_CHP_heat(df_Power_BY, "Base year")


#Calculate energy consumption in all sectors for fuel refinery (FR)
#----------------------------------------------------------------------------------
def calculate_FR_fuel(df_FEC, df_transport_consumption_fuels, df_FuelRefinery):
  
    fuels_FR = df_FuelRefinery.index.tolist() #getting fuel types in Fuel refinery
    df_fuel_secondary_consumption = pd.DataFrame(index=fuels_FR, columns=['Fuel consumption in all sectors'])
    
    #Adding fuel consumption for relevant fuels from transport sector and converting from [TJ] to [PJ]
    for i in fuels_FR:
        if i != 'Biogas':
            df_fuel_secondary_consumption.loc[i] = df_transport_consumption_fuels.loc[i, 'Total transport consumption']/1000
    
    #Adding diesel consumption from Industry sector and converting from [TJ] to [PJ]
    df_fuel_secondary_consumption.loc['Diesel'] += df_FEC['Industry']['Oil Boiler']/1000            
    
    return df_fuel_secondary_consumption
#----------------------------------------------------------------------------------

df_fuel_secondary_consumption_BY = calculate_FR_fuel(df_FEC_BY, df_transport_consumption_fuels_BY, df_FuelRefinery_BY)
df_fuel_secondary_consumption_SC = calculate_FR_fuel(df_FEC_SC, df_transport_consumption_fuels_SC, df_FuelRefinery_SC)

#Calculate total heating demand 
#----------------------------------------------------------------------------------
def calculate_heat_demand(df_FEC, df_fuel_secondary_consumption, df_FuelRefinery):
    #FDHC - Total final DH consumption in sectors
    FDHC = df_FEC.loc['District Heating']['Total by type'] /1000

    #DHFR - DH consumption of fuel refinery
    DHFR = (df_fuel_secondary_consumption['Fuel consumption in all sectors'] * df_FuelRefinery['District Heating']).sum()

    #TDHC - Total gross DH consumption
    TDHC = FDHC + DHFR

    #grid losses (GL)
    GL_DH_coefficient = GL_DH_COEF
    GL_DH = (TDHC / (1 - GL_DH_coefficient)) - TDHC

    #total gross district heating consumption = total heating demand
    GDHC = FDHC + DHFR + GL_DH
    
    # print for validation of model when comparing to STREAM
    # print("FDHC: ", FDHC, "\nDHFR: ", DHFR, "\nGL_DH: ", GL_DH, "\nGDHC: ", GDHC)
    return GDHC
    #----------------------------------------------------------------------------------

total_DH_demand_BY = calculate_heat_demand(df_FEC_BY, df_fuel_secondary_consumption_BY, df_FuelRefinery_BY)
total_DH_demand_SC = calculate_heat_demand(df_FEC_SC, df_fuel_secondary_consumption_SC, df_FuelRefinery_SC)    

    
#Calculate the heating production of District heating boiler
#----------------------------------------------------------------------------------
def calculate_DH_production(df_Heat_shares, total_DH_demand, df_Power):
    total_heat_CHP = df_Power['CHP heat production'].sum()
    total_DH_boiler_production = total_DH_demand - total_heat_CHP
    
    DH_boiler_production = df_Heat_shares * total_DH_boiler_production
    
    return DH_boiler_production
#----------------------------------------------------------------------------------

    
#Calculate total gross electricity consumption 
#----------------------------------------------------------------------------------
def calculate_electricity_demand(df_FEC, df_Heat, df_transport_consumption_fuels, df_FuelRefinery, df_fuel_secondary_consumption, total_DH_demand, df_Power):
    #FECL - Total final electricity consumption in all sectors calculated with consumtion from Electricity consuming technology (Electric appliances, Electric heater, Heatpumps)
    FECL = (df_FEC.iloc[0:2]['Total by type'].sum() + df_FEC.iloc[8]['Total by type'])/1000
    
    #Adding electricity consumption in transport sector (ELTR)
    ELTR = df_transport_consumption_fuels.loc['Electricity']['Total transport consumption']/1000
    FECL += ELTR

    
    #ELFR = Electricity consumption for fuel refinery 
    ELFR = (df_fuel_secondary_consumption['Fuel consumption in all sectors'] * df_FuelRefinery['Electricity']).sum()

    #ELDH = Electricity consumption for DH generation, calculated for the tech types that use electricity as fuel (Heatpump, Electric Boiler)
    ELDH = df_Heat.loc['Heatpump']['Amount']/df_Heat.loc['Heatpump']['Fuel efficiency'] + df_Heat.loc['Electric Boiler']['Amount']/df_Heat.loc['Electric Boiler']['Fuel efficiency']      
    
    #TELC = Total electricity consumption 
    TELC = FECL + ELFR + ELDH
    
    #GL_el - grid losses (GL)
    GL_el = (TELC / (1 - GL_el_COEF)) - TELC

    #GELC - Total gross electricity consumption = total electricity demand
    GELC = FECL + ELFR + ELDH + GL_el
    
    # print for validation of model when comparing to STREAM
    # print("FECL: ", FECL, "\nELFR: ", ELFR, "\nELDH: ", ELDH, "\nGL_el: ", GL_el, "\nGELC : ", GELC)
    return GELC, GL_el, ELTR, ELDH, ELFR
#----------------------------------------------------------------------------------    

#Calculate power amounts 
def calculate_power_production(total_electricity_demand, df_Power_shares):
    df_Power_amount = total_electricity_demand * df_Power_shares
    return df_Power_amount

def final_heatPower_calculate(df_FEC_SC, df_transport_consumption_fuels_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, df_Power_SC):
    #Iteration of calculation for district heating boiler production, based on CHP production and total electricity demand
    #----------------------------------------------------------------------------------
    n_iteration = 0
    tolerance = 0.00000001 #allowed difference between new calculation and last
    iterDiff = 100 
    maxIteration = 1000
    df_Heat_SC['Amount'] = 0
    total_electricity_demand_SC, gridloss_SC, el_transport, el_DH, el_FR = calculate_electricity_demand(df_FEC_SC, df_Heat_SC, df_transport_consumption_fuels_SC, df_FuelRefinery_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, df_Power_SC)
    df_Power_SC['Amount'] = calculate_power_production(total_electricity_demand_SC, df_Power_SC['Shares'])
    
    while iterDiff > tolerance: #Stops when calculation doesn't change much anymore
        total_DH_production = df_Heat_SC['Amount'].sum()
        
        n_iteration += 1
        if n_iteration == maxIteration:
            break
        total_electricity_demand_SC, gridloss_SC, el_transport, el_DH, el_FR = calculate_electricity_demand(df_FEC_SC, df_Heat_SC, df_transport_consumption_fuels_SC, df_FuelRefinery_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, df_Power_SC)
        df_Power_SC['Amount'] = calculate_power_production(total_electricity_demand_SC, df_Power_SC['Shares'])
        df_Power_SC['CHP heat production'], df_Power_SC['Electricity production (condensing or nonCHP)'], df_Power_SC['Fuel consumption in CHP-mode'] = calculate_CHP_heat(df_Power_SC, "Scenario year")
        df_Heat_SC['Amount'] = calculate_DH_production(df_Heat_SC['Shares'], total_DH_demand_SC, df_Power_SC)
        total_DH_production_new = df_Heat_SC['Amount'].sum()
    
        #difference between last calculation and new
        iterDiff = abs(total_DH_production - total_DH_production_new)
    #----------------------------------------------------------------------------------
    
    df_Power_SC['CHP heat production'], df_Power_SC['Electricity production (condensing or nonCHP)'], df_Power_SC['Fuel consumption in CHP-mode'] = calculate_CHP_heat(df_Power_SC, "Scenario year")
    
    #In cases where electricity production is so high that heat production in CHP will cover the total DH demand there should be no DH boiler production
    if df_Power_SC['CHP heat production'].sum() >= total_DH_demand_SC:
        df_Heat_SC['Amount'] = [0 for amount in df_Heat_SC['Amount']]
        #The power production above was calculated with the boiler production of the last iteration, which is negative
        #in this case (negative heat pump use lowers the electricity demand). Recalculate it with the boilers switched off.
        total_electricity_demand_SC, gridloss_SC, el_transport, el_DH, el_FR = calculate_electricity_demand(df_FEC_SC, df_Heat_SC, df_transport_consumption_fuels_SC, df_FuelRefinery_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, df_Power_SC)
        df_Power_SC['Amount'] = calculate_power_production(total_electricity_demand_SC, df_Power_SC['Shares'])
        df_Power_SC['CHP heat production'], df_Power_SC['Electricity production (condensing or nonCHP)'], df_Power_SC['Fuel consumption in CHP-mode'] = calculate_CHP_heat(df_Power_SC, "Scenario year")

    #Final total gross electricity consumption for base year and scenario year
    total_electricity_consumption_SC, gridloss_SC, el_transport, el_DH, el_FR = calculate_electricity_demand(df_FEC_SC, df_Heat_SC,  df_transport_consumption_fuels_SC, df_FuelRefinery_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, df_Power_SC)    
          
    return total_electricity_consumption_SC, gridloss_SC, df_Heat_SC['Amount'], total_electricity_demand_SC, el_transport, el_DH, el_FR
  
#======================================================================================
total_electricity_consumption_BY, gridloss_BY, el_transport_BY, el_DH_BY, el_FR_BY = calculate_electricity_demand(df_FEC_BY, df_Heat_BY, df_transport_consumption_fuels_BY, df_FuelRefinery_BY, df_fuel_secondary_consumption_BY, total_DH_demand_BY, df_Power_BY)
df_Power_BY.loc['Imported electricity', 'Amount'] = total_electricity_consumption_BY - df_Power_BY['Amount'].sum()
df_Power_BY['CHP heat production'], df_Power_BY['Electricity production (condensing or nonCHP)'], df_Power_BY['Fuel consumption in CHP-mode'] = calculate_CHP_heat(df_Power_BY, "Base year")

# final Scenario year calculations
total_electricity_consumption_SC, gridloss_SC, df_Heat_SC['Amount'], total_electricity_demand_SC, el_transport_SC, el_DH_SC, el_FR_SC = final_heatPower_calculate(df_FEC_SC, df_transport_consumption_fuels_SC, df_fuel_secondary_consumption_SC, total_DH_demand_SC, df_Power_SC)


#Fuel Composition, fuel consumption of electricity production and district heat production
#======================================================================================
def fuel_composition(df_Power_fuel, df_Power_elConNoCHP, df_Power_elEff, df_Power_fuelCHP, df_Power_CHPheat, df_Power_heatEffCHP, mode, df_Heat_fuel, df_Heat_amount, df_Heat_fuelEff):
    new_index_mapping = {}
    for old, new in zip(df_Power_fuel.index.values, df_Power_fuel.values):
        new_index_mapping[old] = new
    
    #Fuel consumption for electricity production (Fcep) in non-CHP mode
    Fcep_nonCHP = df_Power_elConNoCHP / df_Power_elEff
    
    #Total fuel consumption electricity
    total_fuel_consumption_el = Fcep_nonCHP + df_Power_fuelCHP
    df_Power_totalFuel = total_fuel_consumption_el
    total_fuel_consumption_el = total_fuel_consumption_el.rename(new_index_mapping) #Change index to fuel names
    
    #Fuel consumption (Fc) for heat production (DH) (allocation) 
    Fc_DH = df_Power_CHPheat / df_Power_heatEffCHP
    Fc_DH = Fc_DH.rename(new_index_mapping) #Change index to fuel names
    #Remove inf values that appear because of division with 0
    is_inf = Fc_DH.isin([np.inf, -np.inf])
    Fc_DH[is_inf] = np.nan
    
    #Fuel consumption (Fc) for electricity production (El) (allocation) 
    Fc_El = pd.Series(index=Fc_DH.index.values)
    for i in range(len(Fc_DH.index.values)): 
        if mode.iloc[i] == 'Condensing' or mode.iloc[i] == 'Fuel free':
            Fc_El.iloc[i] = total_fuel_consumption_el.iloc[i]
        else:
            Fc_El.iloc[i] = total_fuel_consumption_el.iloc[i] - Fc_DH.iloc[i]
            
    new_index_mapping_heat = {}
    for old, new in zip(df_Heat_fuel.index.values, df_Heat_fuel.values):
        new_index_mapping_heat[old] = new
        
    #District heat boiler fuel consumption
    DH_boiler_fuel_consumption = df_Heat_amount / df_Heat_fuelEff
    DH_boiler_fuel_consumption = DH_boiler_fuel_consumption.rename(new_index_mapping_heat)

    return Fc_DH, Fc_El, total_fuel_consumption_el, DH_boiler_fuel_consumption, df_Power_totalFuel

Fc_DH_BY, Fc_El_BY, total_fuel_consumption_BY, DH_boiler_fuel_consumption_BY, df_Power_BY['Total fuel consumption'] = fuel_composition(df_Power_BY['Fuel'], df_Power_BY['Electricity production (condensing or nonCHP)'], df_Power_BY['Electric efficiency (extract unit: condensing mode)'], df_Power_BY['Fuel consumption in CHP-mode'], df_Power_BY['CHP heat production'], df_Power_BY['Heat efficiency in CHP mode (used for allocation)'], df_Power_BY['Mode'], df_Heat_BY['Fuel'], df_Heat_BY['Amount'], df_Heat_BY['Fuel efficiency'])
Fc_DH_SC, Fc_El_SC, total_fuel_consumption_SC, DH_boiler_fuel_consumption_SC, df_Power_SC['Total fuel consumption'] = fuel_composition(df_Power_SC['Fuel'], df_Power_SC['Electricity production (condensing or nonCHP)'], df_Power_SC['Electric efficiency (extract unit: condensing mode)'], df_Power_SC['Fuel consumption in CHP-mode'], df_Power_SC['CHP heat production'], df_Power_SC['Heat efficiency in CHP mode (used for allocation)'], df_Power_SC['Mode'], df_Heat_SC['Fuel'], df_Heat_SC['Amount'], df_Heat_SC['Fuel efficiency'])

#======================================================================================
    
#Resources available and used
#======================================================================================
def resources(df_FEC, df_transport_consumption_fuels, Fc_El, Fc_DH, total_fuel_consumption, df_FuelRefinery, DH_boiler_fuel_consumption, df_fuel_secondary_consumption, gridloss, total_electricity_consumption, total_CHP_heat, total_DHboiler):
    
    #adding biogas consumption from electricity and CHP plants and from district heating boilers
    biogas_power_plants = total_fuel_consumption[total_fuel_consumption.index == 'Biogas'].sum()
    biogas_DH_boilers = DH_boiler_fuel_consumption[DH_boiler_fuel_consumption.index == 'Biogas'].sum()
    df_fuel_secondary_consumption.loc['Biogas'] = biogas_power_plants + biogas_DH_boilers
    #Calculating total primary fuel consumption for Fuel Refinery
    #Inputs that are themselves refinery products (ethanol blended into gasoline) are bought at market price:
    #they count in the fuel costs but not in the energy balance, as in the Excel model
    primary_inputs_FR = [fuel for fuel in df_FuelRefinery.columns if fuel not in df_FuelRefinery.index]
    df_fuel_primary_consumption_FR = df_FuelRefinery[primary_inputs_FR].mul(df_fuel_secondary_consumption['Fuel consumption in all sectors'], axis=0)
    
    #Consumption of primary fuels for fuel refinery
    primary_Fc_FR = df_fuel_primary_consumption_FR.sum()
    
    #Fuel consumption per fuel transport sector, only primary fuels 
    primary_Fc_transport = df_transport_consumption_fuels['Total transport consumption'][['Electricity', 'Natural Gas']] /1000
    
    #Fuel consumption (Fc) for electricity production (El) (allocation) 
    primary_Fc_El = Fc_El[Fc_El.index != 'Biogas']
    primary_Fc_El = primary_Fc_El.groupby(primary_Fc_El.index).sum() #Add duplicate fuels together
    primary_Fc_El['Electricity'] = -1 * (total_electricity_consumption - gridloss)
    
    #Fuel consumption (Fc) for heat production (DH) (allocation) 
    primary_Fc_DH = Fc_DH[Fc_DH.index != 'Biogas']
    primary_Fc_DH = primary_Fc_DH.groupby(primary_Fc_DH.index).sum() #Add duplicate fuels together
    grid_loss_heat = GL_DH_COEF
    primary_Fc_DH['District Heating'] = -1 * (1 -   grid_loss_heat) * total_CHP_heat

    #Fuel consumption (Fc) for district heat boiler production (heat)  
    primary_Fc_heat =  DH_boiler_fuel_consumption[DH_boiler_fuel_consumption.index != 'Biogas']
    primary_Fc_heat = primary_Fc_heat.groupby(primary_Fc_heat.index).sum()
    primary_Fc_heat['District Heating'] = -1 * (1 - grid_loss_heat) * total_DHboiler
    
    #Fuel consumption in sectors Residential, Tertiary, Industry and Agriculture
    #----------------------------------------------------------------------------------
    FEC_fuels = ['Electricity', 'Electricity', 'District Heating', 'Coal', 'Oil', 'Natural Gas', 'Biomass (Straw,woodwaste)', 'Solar Heating', 'Electricity']
    new_index_mapping = {}
    for old, new in zip(df_FEC.index.values, FEC_fuels):
        new_index_mapping[old] = new
    
    primary_Fc_residential = df_FEC['Residential'].rename(new_index_mapping) 
    primary_Fc_residential = primary_Fc_residential.groupby(primary_Fc_residential.index).sum() /1000 #Add duplicate fuels together and convert from TJ to PJ    
    primary_Fc_tertiary = df_FEC['Tertiary'].rename(new_index_mapping)
    primary_Fc_tertiary = primary_Fc_tertiary.groupby(primary_Fc_tertiary.index).sum() /1000 #Add duplicate fuels together and convert from TJ to PJ
    primary_Fc_industry = df_FEC['Industry'].rename(new_index_mapping) 
    primary_Fc_industry = primary_Fc_industry[primary_Fc_industry.index != 'Oil'] #in Industry sector the Oil boiler uses diesel instead of Oil
    primary_Fc_industry = primary_Fc_industry.groupby(primary_Fc_industry.index).sum()  /1000 #Add duplicate fuels together and convert from TJ to PJ
    primary_Fc_agriculture = df_FEC['Agriculture'].rename(new_index_mapping)
    primary_Fc_agriculture = primary_Fc_agriculture.groupby(primary_Fc_agriculture.index).sum()  /1000 #Add duplicate fuels together and convert from TJ to PJ
    
    #----------------------------------------------------------------------------------    
    
    #Fuel consumption (Fc), primary fuels (p)
    #----------------------------------------------------------------------------------
    index_p = primary_Fc_FR.index.to_list() + primary_Fc_transport.index.to_list() + primary_Fc_El.index.to_list() + primary_Fc_heat.index.to_list() + primary_Fc_residential.index.to_list()
    index_p = sorted(list(set(index_p)))
    df_Fc_p = pd.DataFrame(index=index_p)
    df_Fc_p['Fuel Refinery'] = primary_Fc_FR
    df_Fc_p['Transport'] = primary_Fc_transport
    df_Fc_p['Electricity production'] = primary_Fc_El
    df_Fc_p['Heat production'] = primary_Fc_DH
    df_Fc_p['DH boilers'] = primary_Fc_heat
    df_Fc_p['Residential'] = primary_Fc_residential
    df_Fc_p['Tertiary'] = primary_Fc_tertiary
    df_Fc_p['Industry'] = primary_Fc_industry
    df_Fc_p['Agriculture'] = primary_Fc_agriculture
    df_Fc_p['Total primary fuel consumption'] = df_Fc_p.sum(axis=1)
    #----------------------------------------------------------------------------------
    
    #Fuel consumption (Fc), all fuels
    #----------------------------------------------------------------------------------
    df_Fc = pd.DataFrame(index=sorted(list(set(index_p + df_fuel_secondary_consumption.index.to_list()))))
    df_Fc['Fuel Refinery'] = primary_Fc_FR
    df_Fc.update({'Fuel Refinery': -1 * df_fuel_secondary_consumption['Fuel consumption in all sectors']})
    df_Fc['Transport'] = df_transport_consumption_fuels['Total transport consumption'] /1000
    df_Fc['Electricity production'] = Fc_El.groupby(Fc_El.index).sum() 
    df_Fc['Heat production'] = Fc_DH.groupby(Fc_DH.index).sum() 
    df_Fc['DH boilers'] = DH_boiler_fuel_consumption.groupby(DH_boiler_fuel_consumption.index).sum() 
    Fc_residential = df_FEC['Residential'].rename(new_index_mapping)
    df_Fc['Residential'] = Fc_residential.groupby(Fc_residential.index).sum() /1000
    Fc_tertiary =  df_FEC['Tertiary'].rename(new_index_mapping)
    df_Fc['Tertiary'] = Fc_tertiary.groupby(Fc_tertiary .index).sum() /1000
    Fc_industry = df_FEC['Industry'].rename(new_index_mapping).rename({'Oil': 'Diesel'}) #in Industry the oil boiler uses diesel (made by the fuel refinery)
    df_Fc['Industry'] = Fc_industry.groupby(Fc_industry.index).sum() /1000
    Fc_agriculture = df_FEC['Agriculture'].rename(new_index_mapping)
    df_Fc['Agriculture'] =  Fc_agriculture.groupby(Fc_agriculture.index).sum() /1000
    df_Fc.loc['Electricity', 'Electricity production'] = -1 * (total_electricity_consumption - gridloss)
    df_Fc.loc['District Heating','Heat production'] = -1 * (1 - grid_loss_heat) * total_CHP_heat
    df_Fc.loc['District Heating','DH boilers'] = -1 * (1 - grid_loss_heat) * total_DHboiler
    df_Fc['Total fuel consumption'] = df_Fc.sum(axis=1)
    #----------------------------------------------------------------------------------
    
    #Fuel consumption per fuel type
    #----------------------------------------------------------------------------------
    summed_data = {} 
    df_fuelType_consumption = pd.DataFrame()
    for column in df_Fc_p.columns.tolist():
        for fuel_type, fuel_set in FUEL_TYPES.items():
            #Filter the Total primary fuel consumption Series for the keys in the fuel_set
            matching_data = df_Fc_p[column][df_Fc_p[column].index.isin(fuel_set)]
            #Calculate the sum of the matching values
            summed_value = matching_data.sum()
            #Store the summed value in the new dictionary
            summed_data[fuel_type] = summed_value
        df_fuelType_consumption[column] = summed_data
        
    
    #----------------------------------------------------------------------------------
    
    #Fuel consumption per category
    #----------------------------------------------------------------------------------
    summed_data = {} 
    df_fuelCategory_consumption = pd.DataFrame()
    for column in df_fuelType_consumption.columns.tolist():
        for fuel_type, fuel_set in CATEGORIES.items():
            #Filter the Total primary fuel consumption Series for the keys in the fuel_set
            matching_data = df_fuelType_consumption[column][df_fuelType_consumption[column].index.isin(fuel_set)]
            #Calculate the sum of the matching values
            summed_value = matching_data.sum()
            #Store the summed value in the new dictionary
            summed_data[fuel_type] = summed_value
        df_fuelCategory_consumption[column] = summed_data
    #----------------------------------------------------------------------------------
    return df_Fc_p, df_Fc, df_fuelType_consumption, df_fuelCategory_consumption
#======================================================================================
    
df_Fc_p_BY, df_Fc_BY, df_fuelType_consumption_BY, df_fuelCategory_consumption_BY = resources(df_FEC_BY, df_transport_consumption_fuels_BY, Fc_El_BY, Fc_DH_BY, total_fuel_consumption_BY, df_FuelRefinery_BY, DH_boiler_fuel_consumption_BY, df_fuel_secondary_consumption_BY, gridloss_BY, total_electricity_consumption_BY, df_Power_BY['CHP heat production'].sum(), df_Heat_BY['Amount'].sum())  
df_Fc_p_SC, df_Fc_SC, df_fuelType_consumption_SC, df_fuelCategory_consumption_SC = resources(df_FEC_SC, df_transport_consumption_fuels_SC, Fc_El_SC, Fc_DH_SC, total_fuel_consumption_SC, df_FuelRefinery_SC, DH_boiler_fuel_consumption_SC, df_fuel_secondary_consumption_SC, gridloss_SC, total_electricity_consumption_SC, df_Power_SC['CHP heat production'].sum(), df_Heat_SC['Amount'].sum())  

# print for validation of model when comparing to STREAM
# print(df_fuelType_consumption_SC["Total primary fuel consumption"])

#Calculate CO2 emissions for fuels used
#======================================================================================
def calculate_CO2(df_Fc, df_Power, df_TechEff):
    df_fuel_CO2 = pd.DataFrame(index=df_Fc.index.values)

    #Electricity production emissions with and without CCS
    #----------------------------------------------------------------------------------
    #excluding CCS
    new_index_mapping = {}
    for old, new in zip(df_Power.index.values, df_Power['Fuel'].values):
        new_index_mapping[old] = new
    el_CO2 = df_Power['Total fuel consumption'].rename(new_index_mapping)
    el_CO2 =  el_CO2.groupby(el_CO2.index).sum() * df_CO2['Emission Factor']
    
    df_fuel_CO2['Electricity production no CCS'] = el_CO2

    
    #with CCS
    CO2_removal = ASSUMPTIONS['CO2 capture rate of CCS plants']
    fuels = df_Power.index.tolist()
    el_CO2_stored = pd.Series(index=fuels)
    for fuel in fuels:
        if 'CO2-storage' in fuel:
           
            el_CO2_stored[fuel] = df_Power.iloc[fuels.index(fuel)]['Total fuel consumption'] * df_CO2.loc[new_index_mapping.get(fuel)]['Emission Factor'] * CO2_removal
            
        else: 
            el_CO2_stored[fuel] = None
    
    el_CO2_stored = el_CO2_stored.rename(new_index_mapping)
    el_CO2_CCS = el_CO2 - el_CO2_stored.groupby(el_CO2_stored.index).sum()
    
    df_fuel_CO2['Electricity production w. CCS'] = el_CO2_CCS
    #----------------------------------------------------------------------------------
    
    #District heating
    #----------------------------------------------------------------------------------
    df_fuel_CO2['District Heat'] = df_Fc['DH boilers'].mul(df_CO2['Emission Factor'], axis=0) 
    #----------------------------------------------------------------------------------

    #Fuel refinery
    #----------------------------------------------------------------------------------
    df_fuel_CO2['Fuel Refinery'] = df_Fc['Fuel Refinery'].mul(df_CO2['Emission Factor'], axis=0) 
    #----------------------------------------------------------------------------------
    
    #Sectors CO2 emissions
    #----------------------------------------------------------------------------------
    sectors_CO2 = df_Fc[['Residential', 'Tertiary', 'Industry', 'Agriculture', 'Transport']].mul(df_CO2['Emission Factor'], axis=0) 
    
    # df_TechEff.set_index("Tech Type", inplace=True)

    df_fuel_CO2['Residential'] = sectors_CO2.loc[sectors_CO2.index != "Electricity",'Residential'] / df_TechEff.loc[df_TechEff.index != "Electricity",'Residential']
    df_fuel_CO2['Tertiary'] = sectors_CO2.loc[sectors_CO2.index != "Electricity",'Tertiary']  / df_TechEff.loc[df_TechEff.index != "Electricity",'Tertiary']
    industry_eff = df_TechEff.loc[df_TechEff.index != "Electricity",'Industry'].rename({'Oil': 'Diesel'}) #the industry oil boiler uses diesel
    df_fuel_CO2['Industry'] = sectors_CO2.loc[sectors_CO2.index != "Electricity",'Industry']  / industry_eff
    df_fuel_CO2['Agriculture'] = sectors_CO2.loc[sectors_CO2.index != "Electricity",'Agriculture'] / df_TechEff.loc[df_TechEff.index != "Electricity",'Agriculture'] 
    
    df_fuel_CO2['Transport'] = sectors_CO2['Transport']
    #----------------------------------------------------------------------------------
    df_fuel_CO2['Total emissions incl. CCS'] = df_fuel_CO2.iloc[:, 1:].sum(axis=1)
    df_fuel_CO2['Total emissions excl. CCS'] = df_fuel_CO2.iloc[:, 2:-1].sum(axis=1)
    for i in df_fuel_CO2.index:
        if not pd.isna(df_fuel_CO2.loc[i, "Electricity production no CCS"]):
            df_fuel_CO2.loc[i, 'Total emissions excl. CCS'] = df_fuel_CO2.loc[i, 'Total emissions excl. CCS'] + df_fuel_CO2.loc[i, "Electricity production no CCS"]
   
    #CO2 emissions per fuel type
    #----------------------------------------------------------------------------------
    summed_data = {} 
    df_fuelType_CO2 = pd.DataFrame()
    for column in df_fuel_CO2.columns.tolist():
        for fuel_type, fuel_set in FUEL_TYPES.items():
            #Filter the Total primary fuel consumption Series for the keys in the fuel_set
            matching_data = df_fuel_CO2[column][df_fuel_CO2[column].index.isin(fuel_set)]
            #Calculate the sum of the matching values
            summed_value = matching_data.sum()
            #Store the summed value in the new dictionary
            summed_data[fuel_type] = summed_value
        df_fuelType_CO2[column] = summed_data
    # print(df_fuelType_CO2['Total emissions excl. CCS'])    
        
    return df_fuel_CO2, df_fuelType_CO2
#======================================================================================
    
df_fuel_CO2_BY, df_fuelType_CO2_BY = calculate_CO2(df_Fc_BY, df_Power_BY, df_TechEff_BY) 
df_fuel_CO2_SC, df_fuelType_CO2_SC = calculate_CO2(df_Fc_SC, df_Power_SC, df_TechEff_SC)  

# print for validation of model when comparing to STREAM
# print(df_fuelType_CO2_SC['Total emissions incl. CCS'])

#======================================================================================
# Economy calculations

# Discount rate for calculating capital costs
cap_rate = ASSUMPTIONS['Discount rate']
#======================================================================================

# Capital Costs
#----------------------------------------------------------------------------------
def transport_costs(df_Transport, df_TransportWork, df_TransportFuel, df_TransportCosts):
    """Annual costs of the vehicles (investment and maintenance) per fuel [mill. EUR], as in the Excel model:
    transport work [mio. pkm or tkm] x the transport type's share of it x the fuel's share x cost per pkm or tkm."""
    costs = pd.Series(0.0, index=df_TransportFuel.columns)
    for i, transport_type in enumerate(transport_types[:9]): #person transport [0:5], goods transport [5:9]
        if transport_type not in df_TransportCosts.index: #bikes have no costs
            continue
        demand = 'Passenger transport' if i < 5 else 'Freight transport'
        transport_work = df_Transport.loc[demand, 'Amount'] * df_TransportWork.loc[transport_type, 'Allocation of transport work']
        cost_per_km = df_TransportCosts.loc[transport_type].reindex(df_TransportFuel.columns) #empty cells (no cost given) count as 0, as in Excel
        costs += (transport_work * df_TransportFuel.loc[transport_type] * cost_per_km).fillna(0)
    return costs

def capital_costs(df_Power, df_Heat, df_Fc, df_FEC, df_SectorsDetails, df_FuelRefineryCost, df_fuel_secondary_consumption, df_vehicle_costs):
    
    df_capital_costs = pd.DataFrame(index=df_Fc.index)
    
    #Electricity capital costs
    for i in df_Power.index:
        if df_Power.loc[i, 'Full Load Hours'] == 0:
            df_Power.loc[i, 'Capital costs'] = 0
        else:
            df_Power.loc[i, 'Capital costs'] = - npf.pmt(cap_rate, df_Power.loc[i, 'Economic lifetime'], df_Power.loc[i, 'Capital cost [€/MW]'] + df_Power.loc[i, 'Infrastructure costs [€/MW]']) / df_Power.loc[i, 'Full Load Hours'] /3.6
       
    el_capital_costs = df_Power['Amount'].mul(df_Power['Capital costs'])
    new_index_mapping = {}
    for old, new in zip(df_Power.index.values, df_Power['Fuel'].values):
        new_index_mapping[old] = new
    
    el_capital_costs = el_capital_costs.rename(new_index_mapping)
    el_capital_costs = el_capital_costs.groupby(el_capital_costs.index).sum()
    df_capital_costs['Electricity'] = df_capital_costs.index.map(el_capital_costs)
    
    
    #District heat capital costs
    for i in df_Heat.index:
        if df_Heat.loc[i, 'Full Load Hours'] == 0:
            df_Heat.loc[i, 'Capital costs'] = 0
        else:
            df_Heat.loc[i, 'Capital costs'] = - npf.pmt(cap_rate, df_Heat.loc[i, 'Economic lifetime'], df_Heat.loc[i, 'Capital cost [€/MW]'] + df_Heat.loc[i, 'Infrastructure costs [€/MW]']) / df_Heat.loc[i, 'Full Load Hours'] /3.6
       
    dh_capital_costs = df_Heat['Amount'].mul(df_Heat['Capital costs'])
    new_index_mapping = {}
    for old, new in zip(df_Heat.index.values, df_Heat['Fuel'].values):
        new_index_mapping[old] = new
    
    dh_capital_costs = dh_capital_costs.rename(new_index_mapping)
    dh_capital_costs = dh_capital_costs.groupby(dh_capital_costs.index).sum()
    df_capital_costs['District Heat'] = df_capital_costs.index.map(dh_capital_costs)

    # Capital costs of Fuel Refinery [€/GJ output]
    fr_capital_cost = pd.Series(index=df_FuelRefineryCost.index)
    
    for fuel in fr_capital_cost.index:      
        fr_capital_cost.loc[fuel] = -npf.pmt(cap_rate,df_FuelRefineryCost.loc[fuel, 'Economic lifetime [years]'],df_FuelRefineryCost.loc[fuel, 'Investment - Refinery [€/GJ/year]'],0)
    
    fr_capital_costs = fr_capital_cost.mul(df_fuel_secondary_consumption['Fuel consumption in all sectors'])
    
    df_capital_costs['Fuel Refinery'] = df_capital_costs.index.map(fr_capital_costs)
    
    #Sectors, new index names to match fuels
    new_index_mapping = {}
    for old, new in zip(df_SectorsDetails.index.values, df_SectorsDetails['Fuel'].values):
        new_index_mapping[old] = new
            
    # Residential capital costs
    
    #calculate capital costs [€/GJ NEC]
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Residential Price'] == 0:
            df_SectorsDetails.loc[i,'Residential Capital Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Residential Capital Costs'] = - npf.pmt(cap_rate, df_SectorsDetails.loc[i, 'Residential Economic Lifetime'], df_SectorsDetails.loc[i, 'Residential Price']) / df_SectorsDetails.loc[i, 'Residential NEC Generation']
    
    #calculate capital costs [mill. €]
    res_capital_costs = df_FEC['Residential'].mul(df_SectorsDetails['Residential Capital Costs'])/1000
    #change index to fuels and sum same fuels
    res_capital_costs = res_capital_costs.rename(new_index_mapping)
    res_capital_costs = res_capital_costs.groupby(res_capital_costs.index).sum()
    df_capital_costs['Residential'] = df_capital_costs.index.map(res_capital_costs)
    
    # Tertiary capital costs
    
    #calculate capital costs [€/GJ NEC]
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Tertiary Price'] == 0:
            df_SectorsDetails.loc[i,'Tertiary Capital Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Tertiary Capital Costs'] = - npf.pmt(cap_rate, df_SectorsDetails.loc[i, 'Tertiary Economic Lifetime'], df_SectorsDetails.loc[i, 'Tertiary Price']) / df_SectorsDetails.loc[i, 'Tertiary NEC Generation']
    #calculate capital costs [mill. €]    
    ter_capital_costs = df_FEC['Tertiary'].mul(df_SectorsDetails['Tertiary Capital Costs'])/1000
    #change index to fuels and sum same fuels
    ter_capital_costs = ter_capital_costs.rename(new_index_mapping)
    ter_capital_costs = ter_capital_costs.groupby(ter_capital_costs.index).sum()
    df_capital_costs['Tertiary'] = df_capital_costs.index.map(ter_capital_costs)
    
    
    # Industry capital costs
    
    #calculate capital costs [€/GJ NEC]
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Industry Price'] == 0:
            df_SectorsDetails.loc[i,'Industry Capital Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Industry Capital Costs'] = - npf.pmt(cap_rate, df_SectorsDetails.loc[i, 'Industry Economic Lifetime'], df_SectorsDetails.loc[i, 'Industry Price']) / df_SectorsDetails.loc[i, 'Industry NEC Generation']
    #calculate capital costs [mill. €]   
    ind_capital_costs = df_FEC['Industry'].mul(df_SectorsDetails['Industry Capital Costs'])/1000
    #change index to fuels and sum same fuels
    ind_capital_costs = ind_capital_costs.rename(new_index_mapping)
    ind_capital_costs = ind_capital_costs.rename({'Oil':'Diesel'}) #In Industry Diesel is used instead of oil
    ind_capital_costs = ind_capital_costs.groupby(ind_capital_costs.index).sum()
    df_capital_costs['Industry'] = df_capital_costs.index.map(ind_capital_costs)
    
    # Agriculture capital costs
    #calculate capital costs [€/GJ NEC]
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Agriculture Price'] == 0:
            df_SectorsDetails.loc[i,'Agriculture Capital Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Agriculture Capital Costs'] = - npf.pmt(cap_rate, df_SectorsDetails.loc[i, 'Agriculture Economic Lifetime'], df_SectorsDetails.loc[i, 'Agriculture Price']) / df_SectorsDetails.loc[i, 'Agriculture NEC Generation']
    #calculate capital costs [mill. €]    
    agr_capital_costs = df_FEC['Agriculture'].mul(df_SectorsDetails['Agriculture Capital Costs'])/1000
    #change index to fuels and sum same fuels
    agr_capital_costs = agr_capital_costs.rename(new_index_mapping)
    agr_capital_costs = agr_capital_costs.groupby(agr_capital_costs.index).sum()
    df_capital_costs['Agriculture'] = df_capital_costs.index.map(agr_capital_costs)
    
    # Transportation capital costs: the vehicles' investment and maintenance (see transport_costs)
    
    df_capital_costs['Transportation'] = df_capital_costs.index.map(df_vehicle_costs)
    
    return df_capital_costs
#----------------------------------------------------------------------------------
df_capital_costs_BY = capital_costs(df_Power_BY, df_Heat_BY, df_Fc_BY, df_FEC_BY, df_SectorsDetails_BY, df_FuelRefineryCost_BY, df_fuel_secondary_consumption_BY,
                                    transport_costs(df_Transport_BY, df_TransportWork_BY, df_TransportFuel_BY, df_TransportCosts_BY))
df_capital_costs_SC = capital_costs(df_Power_SC, df_Heat_SC, df_Fc_SC, df_FEC_SC, df_SectorsDetails_SC, df_FuelRefineryCost_SC, df_fuel_secondary_consumption_SC,
                                    transport_costs(df_Transport_SC, df_TransportWork_SC, df_TransportFuel_SC, df_TransportCosts_SC))



# Operation and Maintanence Costs
#----------------------------------------------------------------------------------
def om_costs(df_Power, df_Heat, df_Fc, df_FEC, df_SectorsDetails, df_FuelRefineryCost, df_fuel_secondary_consumption):
    
    df_om_costs = pd.DataFrame(index=df_Fc.index)

    #Electricity Operations and Maintanence costs
    for i in df_Power.index:
        if df_Power.loc[i, 'Full Load Hours'] == 0:
            df_Power.loc[i, 'O&M'] = 0
        else:
            df_Power.loc[i, 'O&M'] = (df_Power.loc[i, 'Fixed O&M [€/MW/year]'] / df_Power.loc[i, 'Full Load Hours'] / 3.6) + (df_Power.loc[i, 'Var. O&M [€/MWh]'] / 3.6)
      
    el_om_costs = df_Power['Amount'].mul(df_Power['O&M'])
    new_index_mapping = {}
    for old, new in zip(df_Power.index.values, df_Power['Fuel'].values):
        new_index_mapping[old] = new
    
    el_om_costs = el_om_costs.rename(new_index_mapping)
    el_om_costs = el_om_costs.groupby(el_om_costs.index).sum()
    df_om_costs['Electricity'] = df_om_costs.index.map(el_om_costs)
    
    #District Heat Operations and Maintanence costs
    for i in df_Heat.index:
        if df_Heat.loc[i, 'Full Load Hours'] == 0:
            df_Heat.loc[i, 'O&M'] = 0
        else:
            df_Heat.loc[i, 'O&M'] = (df_Heat.loc[i, 'Fixed O&M [€/MW/year]'] / df_Heat.loc[i, 'Full Load Hours'] / 3.6) + (df_Heat.loc[i, 'Var. O&M [€/MWh]'] / 3.6)
            
    dh_om_costs = df_Heat['Amount'].mul(df_Heat['O&M'])
    new_index_mapping = {}
    for old, new in zip(df_Heat.index.values, df_Heat['Fuel'].values):
        new_index_mapping[old] = new
    
    dh_om_costs = dh_om_costs.rename(new_index_mapping)
    dh_om_costs = dh_om_costs.groupby(dh_om_costs.index).sum()
    df_om_costs['District Heat'] = df_om_costs.index.map(dh_om_costs)
    
    # Fuel Refinery Operations and Maintanence costs    
    fr_om_cost = df_FuelRefineryCost['Transportation Cost [€/GJ/year]'] + df_FuelRefineryCost['Fixed O&M [€/GJ/year]'] + df_FuelRefineryCost['Var. O&M [€/GJ/year]']
    fr_om_costs = fr_om_cost.mul(df_fuel_secondary_consumption['Fuel consumption in all sectors'])
    df_om_costs['Fuel Refinery'] = df_om_costs.index.map(fr_om_costs)
    
    
    #Sectors, new index names to match fuels
    new_index_mapping = {}
    for old, new in zip(df_SectorsDetails.index.values, df_SectorsDetails['Fuel'].values):
        new_index_mapping[old] = new
    
    # Residential Operations and Maintanence costs
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Residential Price'] == 0:
            df_SectorsDetails.loc[i,'Residential O&M Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Residential O&M Costs'] = df_SectorsDetails.loc[i, 'Residential O&M'] / df_SectorsDetails.loc[i, 'Residential NEC Generation']
            
    #calculate o&m costs [mill. €]
    res_om_costs = df_FEC['Residential'].mul(df_SectorsDetails['Residential O&M Costs'])/1000
    #change index to fuels and sum same fuels
    res_om_costs = res_om_costs.rename(new_index_mapping)
    res_om_costs = res_om_costs.groupby(res_om_costs.index).sum()
    df_om_costs['Residential'] = df_om_costs.index.map(res_om_costs)
    
    
    # Tertiary Operations and Maintanence costs
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Tertiary Price'] == 0:
            df_SectorsDetails.loc[i,'Tertiary O&M Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Tertiary O&M Costs'] = df_SectorsDetails.loc[i, 'Tertiary O&M'] / df_SectorsDetails.loc[i, 'Tertiary NEC Generation']
            
    #calculate o&m costs [mill. €]
    ter_om_costs = df_FEC['Tertiary'].mul(df_SectorsDetails['Tertiary O&M Costs'])/1000
    #change index to fuels and sum same fuels
    ter_om_costs = ter_om_costs.rename(new_index_mapping)
    ter_om_costs = ter_om_costs.groupby(ter_om_costs.index).sum()
    df_om_costs['Tertiary'] = df_om_costs.index.map(ter_om_costs)
    
    
    # Industry Operations and Maintanence costs
    # Industry also has variable O&M [€/GJ] (as in the Excel model); data files from before it was added count it as 0
    if 'Industry Variable O&M [€/GJ]' not in df_SectorsDetails:
        df_SectorsDetails['Industry Variable O&M [€/GJ]'] = 0
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Industry Price'] == 0:
            df_SectorsDetails.loc[i,'Industry O&M Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Industry O&M Costs'] = df_SectorsDetails.loc[i, 'Industry O&M'] / df_SectorsDetails.loc[i, 'Industry NEC Generation'] + df_SectorsDetails.loc[i, 'Industry Variable O&M [€/GJ]']
            
    #calculate o&m costs [mill. €]
    ind_om_costs = df_FEC['Industry'].mul(df_SectorsDetails['Industry O&M Costs'])/1000
    #change index to fuels and sum same fuels
    ind_om_costs = ind_om_costs.rename(new_index_mapping)
    ind_om_costs = ind_om_costs.rename({'Oil':'Diesel'}) #In Industry Diesel is used instead of oil
    ind_om_costs = ind_om_costs.groupby(ind_om_costs.index).sum()
    df_om_costs['Industry'] = df_om_costs.index.map(ind_om_costs)
    
    # Agriculture Operations and Maintanence costs
    for i in df_SectorsDetails.index:
        if df_SectorsDetails.loc[i, 'Agriculture Price'] == 0:
            df_SectorsDetails.loc[i,'Agriculture O&M Costs'] = 0
        else:
            df_SectorsDetails.loc[i,'Agriculture O&M Costs'] = df_SectorsDetails.loc[i, 'Agriculture O&M'] / df_SectorsDetails.loc[i, 'Agriculture NEC Generation']
            
    #calculate o&m costs [mill. €]
    agr_om_costs = df_FEC['Agriculture'].mul(df_SectorsDetails['Agriculture O&M Costs'])/1000
    #change index to fuels and sum same fuels
    agr_om_costs = agr_om_costs.rename(new_index_mapping)
    agr_om_costs = agr_om_costs.groupby(agr_om_costs.index).sum()
    df_om_costs['Agriculture'] = df_om_costs.index.map(agr_om_costs)
    
    # Transportation Operations and Maintanence costs
    
    df_om_costs['Transportation'] = np.nan
    
    
    return df_om_costs

df_om_costs_BY = om_costs(df_Power_BY, df_Heat_BY, df_Fc_BY, df_FEC_BY, df_SectorsDetails_BY, df_FuelRefineryCost_BY, df_fuel_secondary_consumption_BY)
df_om_costs_SC = om_costs(df_Power_SC, df_Heat_SC, df_Fc_SC, df_FEC_SC, df_SectorsDetails_SC, df_FuelRefineryCost_SC, df_fuel_secondary_consumption_SC)

    
#----------------------------------------------------------------------------------

# Fuel Costs
#----------------------------------------------------------------------------------
def fuel_costs(df_FuelPrice, df_FuelRefinery, df_FuelRefineryCost, df_Fc):
          
    # Refined price [€/GJ] (total output fuel cost [€/GJ output]) for Fuel Refinery
        # is the sum of total input fuel costs, capital costs, operation costs (fixed), & operation costs (variable)
    
    # Fuel input costs of Fuel Refinery
    df_fuel_input_cost_FR = df_FuelRefinery.mul(df_FuelPrice['Market price [€/GJ]'], axis=1)
    
    # Total input fuel costs of Fuel Refinery [€/GJ output]    
    total_fuel_input_cost_FR = df_fuel_input_cost_FR.sum(axis=1)

    # Capital costs of Fuel Refinery [€/GJ output]
    capital_cost_FR = pd.DataFrame(index=df_FuelRefineryCost.index)
    
    for fuel in capital_cost_FR.index:      
        capital_cost_FR.loc[fuel, 'Capital costs [€/GJ output]]'] = -npf.pmt(cap_rate,df_FuelRefineryCost.loc[fuel, 'Economic lifetime [years]'],df_FuelRefineryCost.loc[fuel, 'Investment - Refinery [€/GJ/year]'],0)
    
    # Operation costs, fixed of Fuel Refinery [€/GJ output]
    
    operation_costFix_FR = df_FuelRefineryCost['Transportation Cost [€/GJ/year]'] + df_FuelRefineryCost['Fixed O&M [€/GJ/year]']
    
    # Operation costs, variable of Fuel Refinery [€/GJ output]
    operation_costVar_FR = df_FuelRefineryCost['Var. O&M [€/GJ/year]']

    # Refined price [€/GJ] (total output fuel cost [€/GJ output]) of Fuel Refinery
    total_fuel_output_cost_FR = pd.DataFrame(index=df_FuelRefineryCost.index)
    total_fuel_output_cost_FR['Total output fuel cost [€/GJ output]]'] = total_fuel_input_cost_FR + capital_cost_FR['Capital costs [€/GJ output]]'] + operation_costFix_FR + operation_costVar_FR 
    
    # Total fuel price for all fuel types [€/GJ]
    df_FuelPrice['Final price [€/GJ]'] = df_FuelPrice['Market price [€/GJ]']
    df_FuelPrice.loc[df_FuelPrice.index.isin(total_fuel_output_cost_FR.index), 'Final price [€/GJ]'] = total_fuel_output_cost_FR['Total output fuel cost [€/GJ output]]']

    # Total fuel cost for all fuel types [€]
    df_fuel_costs = df_Fc.mul(df_FuelPrice['Final price [€/GJ]'], axis = 0).drop(columns=['Fuel Refinery','Total fuel consumption'])
    df_fuel_costs['Fuel input cost for fuel refinery'] = -total_fuel_input_cost_FR.mul(df_Fc['Fuel Refinery'])
    df_fuel_costs['Fuel output cost for fuel refinery'] = total_fuel_output_cost_FR['Total output fuel cost [€/GJ output]]'].mul(df_Fc['Fuel Refinery'])
    
    # Total Fuel refinery fuel costs is the value of the fuel input minus the value of the fuel out (likely a negative result)
    df_fuel_costs['Fuel Refinery'] = df_fuel_costs['Fuel input cost for fuel refinery'] + df_fuel_costs['Fuel output cost for fuel refinery']
    
    # Costs of fuels for CHP heat and DH boilers is summed as total fuel costs for District Heat
    # (fill_value=0: a fuel used only in boilers, e.g. oil or heat imports, has no CHP value; '+' would make its cost empty)
    df_fuel_costs['District Heat'] = df_fuel_costs['DH boilers'].add(df_fuel_costs['Heat production'], fill_value=0)
    
    # Removing irrelavant columns and reordering columns for consistency
    df_fuel_costs = df_fuel_costs.drop(columns=['DH boilers', 'Heat production', 'Fuel input cost for fuel refinery', 'Fuel output cost for fuel refinery'])
    df_fuel_costs = df_fuel_costs.rename(columns={'Electricity production' : 'Electricity'})
    df_fuel_costs = df_fuel_costs[['Electricity', 'District Heat', 'Fuel Refinery', 'Residential', 'Tertiary', 'Industry', 'Agriculture', 'Transport']]
    
    
    # summed_data = {} 
    # df_fuelType_costs = pd.DataFrame()
    # for column in df_fuel_costs.columns.tolist():
    #     if column != 'Fuel input cost for fuel refinery':
    #         for fuel_type, fuel_set in FUEL_TYPES_PRIMARY.items():
    #             #Filter the Total primary fuel consumption Series for the keys in the fuel_set
    #             matching_data = df_fuel_costs[column][df_fuel_costs[column].index.isin(fuel_set)]
    #             #Calculate the sum of the matching values
    #             summed_value = matching_data.sum()
    #             #Store the summed value in the new dictionary
    #             summed_data[fuel_type] = summed_value
    #         df_fuelType_costs[column] = summed_data
    #     else:
    #         for fuel_type, fuel_set in FUEL_TYPES.items():
    #             #Filter the Total primary fuel consumption Series for the keys in the fuel_set
    #             matching_data = df_fuel_costs[column][df_fuel_costs[column].index.isin(fuel_set)]
    #             #Calculate the sum of the matching values
    #             summed_value = matching_data.sum()
    #             #Store the summed value in the new dictionary
    #             summed_data[fuel_type] = summed_value
    #         df_fuelType_costs[column] = summed_data
            
    # df_fuelType_costs['Electricity and CHP'] = df_fuelType_costs['Electricity production'] +  df_fuelType_costs['Heat production']
    
    # df_fuelType_costs = df_fuelType_costs.drop(columns=['Fuel output cost for fuel refinery','Electricity production','Heat production'],axis=1)
    
    return df_FuelPrice, df_fuel_costs
    
    
    
#----------------------------------------------------------------------------------
df_FuelPrice_BY, df_fuel_costs_BY = fuel_costs(df_FuelPrice_BY, df_FuelRefinery_BY, df_FuelRefineryCost_BY, df_Fc_BY)
df_FuelPrice_SC, df_fuel_costs_SC = fuel_costs(df_FuelPrice_SC, df_FuelRefinery_SC, df_FuelRefineryCost_SC, df_Fc_SC)



# CO2 Costs
#----------------------------------------------------------------------------------
def co2_costs(co2_price, df_fuel_CO2):
    # Cost of CO2 emissions is based on the amount of CO2 emitted and a set price pr ton [€/t CO2]
    df_co2_costs = df_fuel_CO2.mul(co2_price)
    df_co2_costs[df_co2_costs <= 0] = np.nan # CO2 costs cannot be negative
    # Electricity pays for the CO2 that is emitted, i.e. after CO2 capture and storage (CCS), as in the Excel model
    df_co2_costs = df_co2_costs.drop(['Electricity production no CCS', 'Total emissions incl. CCS', 'Total emissions excl. CCS'], axis=1) #remove irrelevant columns
    df_co2_costs = df_co2_costs.rename(columns={'Electricity production w. CCS' : 'Electricity'}) # Rename electricity column for simplicity and consistency
    # The CO2 of refined fuels (diesel, gasoline, ...) is paid where they are used (transport, industry, ...), as in the Excel model.
    # The fuel refinery itself pays nothing, otherwise the same CO2 would be paid twice (as crude oil in the refinery and as diesel).
    df_co2_costs['Fuel Refinery'] = np.nan
    return df_co2_costs

#----------------------------------------------------------------------------------
co2_price_BY = ASSUMPTIONS['CO2 price, base year']

df_co2_costs_BY = co2_costs(co2_price_BY, df_fuel_CO2_BY)
df_co2_costs_SC = co2_costs(co2_price_SC, df_fuel_CO2_SC)

#======================================================================================