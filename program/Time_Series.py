
"""
Extension of pySTREAM

@author: Ida Cecilie Hoielt
"""
import pandas as pd
import numpy as np
import os

#Only constants are imported from the backend. Scenario results are passed to run_timeSeries() as arguments,
#so the time series always uses the current scenario (names imported with * would keep their startup values)
from Backend import GL_el_COEF, DATA_DIR, ASSUMPTIONS


IntermittentRenewables = ['Wind, offshore','Wind, onshore','Photo voltaic','Wave power','Hydro']



#Load Energy profiles data
#==============================================================================

ProfilesData = os.path.join(DATA_DIR, 'Energy profiles.xlsx')

df_profile_ElConsumption = pd.read_excel(ProfilesData, sheet_name='Consumption')
df_profile_HeatConsumption = pd.read_excel(ProfilesData, sheet_name='Heat consumption')
df_profile_Wind = pd.read_excel(ProfilesData, sheet_name='Wind')
df_profile_Solar = pd.read_excel(ProfilesData, sheet_name='Solar')
df_profile_Hydro = pd.read_excel(ProfilesData, sheet_name='Hydro')
df_profile_Wave = pd.read_excel(ProfilesData, sheet_name='Wave')
df_nightHours = pd.read_excel(ProfilesData, sheet_name='Night hours')

#==============================================================================


#Electricity consumption
#==============================================================================
def el_consumption(df_FEC_SC, total_electricity_demand_SC, el_transport_SC, el_DH_SC, el_FR_SC):
    el_demand_Heatpumps = (df_FEC_SC.loc['Heatpumps', 'Total by type']/(1-GL_el_COEF))/1000
    el_demand_ElHeater = (df_FEC_SC.loc['Electric Heater','Total by type']/(1-GL_el_COEF))/1000
    general_el_consumption = +(total_electricity_demand_SC - (el_transport_SC/(1-GL_el_COEF)) - el_demand_Heatpumps - el_demand_ElHeater - el_DH_SC - el_FR_SC/(1-GL_el_COEF))
    #general_el_consumption -= df_Power_SC.loc['Imported electricity', 'Amount']
    # calculate general electricity consumption in MWh
    general_el_consumption_MWh = (general_el_consumption * 10**6 ) / 3.6
    #total_profile full load hours is the sum of all profile entries
    total_profile_ElConsumption = df_profile_ElConsumption['Consumption profile'].sum(axis=0) 
    df_hour_ElConsumption = (df_profile_ElConsumption['Consumption profile'] * (general_el_consumption_MWh/total_profile_ElConsumption)).rename('Consumption')
    
    
    #Heat pumps
    total_profile_HeatConsumption = df_profile_HeatConsumption['Heat consumption profile'].sum(axis=0)
    df_hour_ElConsumption_heatpump = df_profile_HeatConsumption['Heat consumption profile'] / total_profile_HeatConsumption * (el_demand_Heatpumps * 10**6 / 3.6)

    #Electric heaters
    df_hour_ElConsumption_elheater = df_profile_HeatConsumption['Heat consumption profile'] / total_profile_HeatConsumption * (el_demand_ElHeater * 10**6 / 3.6)

    #Heat pumps and electric boilers in district heating. They are subtracted from the general consumption above, as in
    #STREAM's export to its hourly model, which runs them with the heat production. pySTREAM has no hourly heat production,
    #so their electricity follows the heat consumption profile (otherwise it would be missing from the hourly consumption).
    df_hour_ElConsumption_DH = df_profile_HeatConsumption['Heat consumption profile'] / total_profile_HeatConsumption * (el_DH_SC * 10**6 / 3.6)

    df_hour_ElConsumption = df_hour_ElConsumption + df_hour_ElConsumption_heatpump + df_hour_ElConsumption_elheater + df_hour_ElConsumption_DH


    return df_hour_ElConsumption.rename('Consumption')

#==============================================================================


#Wind power
#==============================================================================
def wind_power(windOffPower_amount, windOnPower_amount):

    #Each wind type (onshore/offshore) follows two profiles: profile 1 gets share_wind of the production, profile 2 the rest
    share_wind = ASSUMPTIONS['Time series: wind production following profile 1']

    #Offshore

    #total_profile full load hours is the sum of all profile entries
    total_profile_WindOff1 = df_profile_Wind['Offshore 1 profile'].sum(axis=0)
    total_profile_WindOff2 = df_profile_Wind['Offshore 2 profile'].sum(axis=0)

    installed_cap_WindOff1 = (windOffPower_amount / 3.6 * 10**6) * share_wind / total_profile_WindOff1
    installed_cap_WindOff2 = (windOffPower_amount / 3.6 * 10**6) * (1 - share_wind) / total_profile_WindOff2

    hour_WindOff1 = (df_profile_Wind['Offshore 1 profile']*installed_cap_WindOff1).to_numpy()
    hour_WindOff2 = (df_profile_Wind['Offshore 2 profile']*installed_cap_WindOff2).to_numpy()

    #Onshore

    #total_profile full load hours is the sum of all profile entries
    total_profile_WindOn1 = df_profile_Wind['Onshore 1 profile'].sum(axis=0)
    total_profile_WindOn2 = df_profile_Wind['Onshore 2 profile'].sum(axis=0)

    installed_cap_WindOn1 = (windOnPower_amount / 3.6 * 10**6) * share_wind / total_profile_WindOn1
    installed_cap_WindOn2 = (windOnPower_amount / 3.6 * 10**6) * (1 - share_wind) / total_profile_WindOn2

    hour_WindOn1 = (df_profile_Wind['Onshore 1 profile']*installed_cap_WindOn1).to_numpy()
    hour_WindOn2 = (df_profile_Wind['Onshore 2 profile']*installed_cap_WindOn2).to_numpy()

    wind_data =  {
        'Offshore 1 [MWh per hour]': hour_WindOff1,
        'Offshore 2 [MWh per hour]': hour_WindOff2,
        'Onshore 1 [MWh per hour]': hour_WindOn1,
        'Onshore 2 [MWh per hour]': hour_WindOn2
    }

    df_hour_Wind = pd.DataFrame(wind_data)

    df_hour_Wind['Total Wind'] = df_hour_Wind.sum(axis=1)
    df_hour_Wind['Total Offshore'] = df_hour_Wind['Offshore 1 [MWh per hour]'] + df_hour_Wind['Offshore 2 [MWh per hour]']
    df_hour_Wind['Total Onshore'] = df_hour_Wind['Onshore 1 [MWh per hour]'] + df_hour_Wind['Onshore 2 [MWh per hour]']

    #Required installed capacity
    required_cap_Off = installed_cap_WindOff1 + installed_cap_WindOff2 
    required_cap_On = installed_cap_WindOn1 + installed_cap_WindOn2

    return df_hour_Wind['Total Offshore'], df_hour_Wind['Total Onshore'], required_cap_Off, required_cap_On

#==============================================================================


#Solar power
#==============================================================================
def solar_power(solarPower_amount):
    total_SolarPower_productionMWh = solarPower_amount * 10**6 / 3.6

    #total_profile full load hours is the sum of all profile entries
    total_profile_Solar = df_profile_Solar['Solar profile'].sum(axis=0)

    df_hour_solarEl = (df_profile_Solar['Solar profile'] / total_profile_Solar * total_SolarPower_productionMWh).rename('Solar Power [MWh per hour]')

    return df_hour_solarEl

#==============================================================================


#Wave power
#==============================================================================
def wave_power(wavePower_amount):
    total_Wave_productionMWh = wavePower_amount * 10**6 / 3.6

    #total_profile full load hours is the sum of all profile entries
    total_profile_Wave = df_profile_Wave['Wave profile'].sum(axis=0)

    df_hour_wave = (df_profile_Wave['Wave profile'] / total_profile_Wave * total_Wave_productionMWh).rename('Wave [MWh per hour]')


    return df_hour_wave

#==============================================================================

#Hydro power
#==============================================================================
def hydro_power(hydroPower_amount):

    total_Hydro_productionMWh = hydroPower_amount * 10**6 / 3.6

    #total_profile full load hours is the sum of all profile entries
    total_profile_Hydro = df_profile_Hydro['Hydro profile'].sum(axis=0)

    df_hour_hydro = (df_profile_Hydro['Hydro profile'] / total_profile_Hydro * total_Hydro_productionMWh).rename('Hydro [MWh per hour]')

    return df_hour_hydro

#==============================================================================

#Solar heat
#==============================================================================
def solar_heat(solarHeat_amount):
    total_SolarHeat_productionMWh = solarHeat_amount * 10**6 / 3.6

    #total_profile full load hours is the sum of all profile entries
    total_profile_Solar = df_profile_Solar['Solar profile'].sum(axis=0)

    df_hour_solarHeat = (df_profile_Solar['Solar profile'] / total_profile_Solar * total_SolarHeat_productionMWh).rename('Solar Heat [MWh per hour]')

    return df_hour_solarHeat

#df_hour['Solar heat'] = solar_power(df_Heat_SC.loc['Solar heat', 'Amount'])
#==============================================================================

#Consumption balance after intermittent renewables (electricity consumption minus wind, hydro, solar, wave)
#==============================================================================
def consumption_balance(df_hour):
    intermittentRenewables_hour = ['Wind, offshore','Wind, onshore','Photo voltaic','Wave power','Hydro']
    intermittentRenewables_power = df_hour[intermittentRenewables_hour].sum(axis=1)
    balance = df_hour['Consumption'] - intermittentRenewables_power

    return intermittentRenewables_power, balance

#==============================================================================


def consumption_flexible(el_transport, el_FR, consumption_after_renew):

    share_FR_veryFlexible = ASSUMPTIONS['Time series: fuel refinery electricity when best for the system']
    share_EV_veryFlexible = ASSUMPTIONS['Time series: EV charging when best for the system']
    share_night_consumption = ASSUMPTIONS['Time series: EV charging at night'] #night hours are marked in Energy profiles.xlsx

    increased_el_consumption_EV = el_transport/(1-GL_el_COEF) * 10**6 / 3.6
    increased_el_consumption_FR = el_FR/(1-GL_el_COEF) * 10**6 / 3.6

    #Intersection: Very flexible demand is added to those hours with a consumption level below the intersection value. The intersection must be lower than the maximum hourly consumption and large enough to prevent an increase of the peak consumption (observe 'Intersection Check'). If the error 'Div#/0' appears in the check field, the intersection is too low.
    intersection = 0

    #unflexible electricity consumption for Fuel Refinery (MW): when the refinery uses electricity, the very flexible share is
    #used when best for the system (below) and the rest evenly over the year. When it produces electricity on balance
    #(negative), all of it is spread evenly over the year.
    if increased_el_consumption_FR > 0:
        unflexible_FR = increased_el_consumption_FR * (1-share_FR_veryFlexible) / 8760
    else:
        unflexible_FR = increased_el_consumption_FR / 8760

    #period for very flexible consumption
    period_flexible = pd.Series(0, index = consumption_after_renew.index)
    #add the unflexible consumption: electric vehicles not charged when best for the system or at night, and the fuel refinery
    consumption_after_renew = consumption_after_renew + ((increased_el_consumption_EV*(1-share_EV_veryFlexible-share_night_consumption))/8760+unflexible_FR)
    period_flexible[consumption_after_renew < intersection] = 1

    #hours of increase for very flexible
    total_hours_flexible = period_flexible.sum()
    if total_hours_flexible == 0:
        #No hours with surplus renewable power (e.g. no wind or solar): spread the very flexible consumption evenly over the year
        #instead of dividing by zero, which made all hourly values NaN and the capacity calculation below run forever
        period_flexible[:] = 1
        total_hours_flexible = len(period_flexible)
    veryFlexible_EV = share_EV_veryFlexible * increased_el_consumption_EV
    veryFlexible_FR = share_FR_veryFlexible * increased_el_consumption_FR
    veryFlexible_FR_MW = 0
    if veryFlexible_FR > 0:
        veryFlexible_FR_MW = veryFlexible_FR / total_hours_flexible


    

    #total increase per hour of very flexible EV and FR
    total_increase = veryFlexible_EV / total_hours_flexible + veryFlexible_FR_MW

    #Electricity consumption, flexible (when best for system)
    el_consumption_FlexibleBest = total_increase * period_flexible 

    #hours of increase of night consumption
    total_night_hours = df_nightHours['Night hours'].sum()

    #Electricity consumption, flexible (night)
    total_night_consumption = share_night_consumption * increased_el_consumption_EV
    el_consumption_FlexibleNight = (total_night_consumption / total_night_hours) * df_nightHours['Night hours']


    #Total consumption, flexible and unflexible
    el_consumption_total = consumption_after_renew + el_consumption_FlexibleBest + el_consumption_FlexibleNight


    return el_consumption_FlexibleBest, el_consumption_FlexibleNight, el_consumption_total



def production_flexible(power_amount, consumption, power_shares, df_hour, df_fuel, df_FuelPrice):
    flexibleTechnologies = [index for index in power_amount.index if index not in IntermittentRenewables]
    intermittentRenewables_share = power_shares[IntermittentRenewables].sum()
    flexible_technologies_shares = power_shares[flexibleTechnologies]
    flexible_technologies_share = flexible_technologies_shares.sum()

    export = pd.Series(0.0, index=consumption.index, name='Exported electricity')
    export[consumption < 0] +=  consumption * -1

    consumption_pos = pd.DataFrame(0.0, index=consumption.index, columns=['Consumption'] )
    
    consumption_pos.loc[consumption > 0, 'Consumption'] = consumption
    consumption_pos_sum = consumption_pos.sum()

    share_MWh = (power_shares[flexibleTechnologies] / flexible_technologies_share) * consumption_pos_sum
    
    #Technologies used for flexible production, inlcuding share, amount and priorities
    technologies_used = pd.DataFrame(index=flexible_technologies_shares[flexible_technologies_shares > 0].index, columns=['Amount', 'Shares', 'Fuel', 'Priority','Capacity', 'Hours', 'Avg'])
    


    #Duration curve
    sorted_consumption = pd.DataFrame(consumption_pos['Consumption'].sort_values(ascending=False)) 

    for tech in flexibleTechnologies:
        sorted_consumption[tech] = 0.0
        consumption_pos[tech] = 0.0

    for tech in technologies_used.index:
        technologies_used.loc[tech, 'Amount'] = power_amount.loc[tech] * 10**6 /3.6
        technologies_used.loc[tech, 'Shares'] = power_shares.loc[tech]
        technologies_used.loc[tech, 'Fuel'] = df_fuel.loc[tech]
        technologies_used.loc[tech, 'Priority'] = df_FuelPrice.loc[df_fuel.loc[tech],'Final price [€/GJ]']
        technologies_used.loc[tech, 'Capacity'] = 0
        technologies_used.loc[tech, 'Hours'] = 0
        technologies_used.loc[tech, 'Avg'] = 0

    technologies_used = technologies_used.sort_values(by='Priority', ascending=True)


    # if power_shares['Imported electricity'] < 0:
    #     technologies_used.loc['Exported electricity','Amount'] = -1*power_amount.loc['Imported electricity'] * 10**6 /3.6
    #     technologies_used.loc['Exported electricity','Shares'] = -1*power_shares.loc['Imported electricity']
    #     technologies_used.loc['Exported electricity','Fuel'] = 'Exported electricity'
    #     technologies_used.loc['Exported electricity','Priority'] = None
    #     technologies_used.loc['Exported electricity','Capacity'] = 0
    #     technologies_used.loc['Exported electricity','Hours'] = 0
    #     technologies_used.loc['Exported electricity','Avg'] = 0
    #     consumption_pos['Exported electricity'] = 0.0
    #     sorted_consumption['Exported electricity'] = 0.0

    
    digits = 1
    step = 1
    print('Calculating capacities for flexible technologies...')
    progress = 0
    for tech in technologies_used.itertuples():
        print(int((progress / len(technologies_used.index)) * 100) , '%')
        progress += 1
        cap = 0
        while consumption_pos[tech.Index].sum() < tech.Amount:
            hours_with_demand = consumption_pos['Consumption'] > 0
            if not hours_with_demand.any():
                break #no demand left to cover, so adding capacity would never end the loop
            cap += step
            consumption_pos.loc[hours_with_demand, tech.Index] +=  step
            consumption_pos.loc[hours_with_demand, 'Consumption'] -= step
            if consumption_pos['Consumption'].sum() < 0:
                break
        technologies_used.loc[tech.Index, 'Capacity'] = round(cap, ndigits=digits) 
   
    print(int(100), '%')
    

    # Initialize contribution dictionary
    contribution = {tech: 0 for tech in technologies_used.index}

    for i in sorted_consumption.index:
         remaining_consumption = sorted_consumption.loc[i, 'Consumption']
         for tech in technologies_used.itertuples():
             #tech_capacity = (tech['Shares'] / flexible_technologies_share) * consumption_pos_sum
             if contribution[tech.Index] < tech.Amount:
                 allocation = min(remaining_consumption, tech.Capacity)
                 contribution[tech.Index] += allocation
                 sorted_consumption.loc[i, tech.Index] = allocation
                 remaining_consumption -= allocation
             if remaining_consumption <= 0:
                 break

    for tech in technologies_used.index:
        technologies_used.loc[tech, 'Hours'] = sorted_consumption.loc[sorted_consumption[tech] > 0, tech].count()
        hours = technologies_used.loc[tech, 'Hours']
        technologies_used.loc[tech, 'Avg'] = technologies_used.loc[tech, 'Amount'] / hours if hours > 0 else 0
   


    total_balance = consumption - sorted_consumption.iloc[:,1:].sort_index().sum(axis=1) + export
    
    sorted_consumption.loc[total_balance > 0, 'Imported electricity'] += total_balance[total_balance > 0]
    

    #sorted_consumption['Exported electricity'] += -1* total_balance[total_balance < 0]
    

    df_hour = pd.concat([df_hour, sorted_consumption.sort_index(), export, total_balance.rename("total")-sorted_consumption['Imported electricity']], axis=1)
    
    return df_hour, technologies_used


def run_timeSeries(df_Power_SC, df_FuelPrice_SC, df_FEC_SC, total_electricity_demand_SC, el_transport_SC, el_DH_SC, el_FR_SC):
    df_hour = pd.DataFrame(el_consumption(df_FEC_SC, total_electricity_demand_SC, el_transport_SC, el_DH_SC, el_FR_SC))
    df_hour['Wind, offshore'], df_hour['Wind, onshore'], required_cap_OffWind, required_cap_OnWind = wind_power(df_Power_SC.loc['Wind, offshore', 'Amount'],df_Power_SC.loc['Wind, onshore', 'Amount'])
    df_hour['Photo voltaic'] = solar_power(df_Power_SC.loc['Photo voltaic', 'Amount'])
    df_hour['Wave power'] = wave_power(df_Power_SC.loc['Wave power', 'Amount'] )
    df_hour['Hydro'] = hydro_power(df_Power_SC.loc['Hydro', 'Amount'])
    df_hour['Renewables power'], df_hour['Consumption after renewables'] = consumption_balance(df_hour)
    df_hour['Flexible consumption when best for system'], df_hour['Flexible consumption night'], df_hour['Consumption after renewables (unflexible and flexible)'] = consumption_flexible(el_transport_SC, el_FR_SC, df_hour['Consumption after renewables'])
    df_hour, technologies_used = production_flexible(df_Power_SC['Amount'], df_hour['Consumption after renewables (unflexible and flexible)'], df_Power_SC['Shares'], df_hour, df_Power_SC['Fuel'], df_FuelPrice_SC)
    
    return df_hour, technologies_used

