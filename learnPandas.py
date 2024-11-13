import pandas as pd
df = pd.read_excel('pba.xlsx')

cols = [
    "Device Name",
    "Starting time",
    "End Time",
    "Duration (m)",
    "Description of running status word",
    "error code",
    "Fault description",
    "Lost power generation (kWh)",
]

new_df = df[cols][df['Description of running status word'].isin(['Fault stop',
                                                           'Tower base stop',
                                                           'Tower base emergency stop',
                                                           'Service mode',
                                                           'Periodic service stop',
                                                           'HMI stop',
                                                           'Nacelle stop',])]

new_df.to_excel('new_pba.xlsx', index = False)


