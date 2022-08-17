import pandas as pd
import numpy as np

def table1(borough):

    df = pd.read_excel("accessible-pedestrian-signals.xlsx")

    df = df[df['Borough'].str.contains(borough)]
    count = df['Borough'].value_counts()[borough]
    data = {'Borough': [borough], 'Total Entries': [count]}

    bdata = pd.DataFrame(data)
    #print(df)
    year = df.groupby(pd.Grouper(key='Date Installed', freq='Y')).size().reset_index(name='Year Count')
    df2 = pd.DataFrame(year)
    df2['Date Installed'] =df2['Date Installed'].dt.strftime('%Y')

    month = df.groupby(pd.Grouper(key='Date Installed', freq='M')).size().reset_index(name='Month Count')
    df3 = pd.DataFrame(month)
    df3['Date Installed'] =df3['Date Installed'].dt.strftime('%Y-%m')

    result = pd.concat([bdata, df2, df3], ignore_index=True)
    result = result.sort_values(by=['Borough', 'Date Installed'], ascending=[True, True])

    return result

def table2():

    df = pd.read_excel("accessible-pedestrian-signals.xlsx")

    year = df.groupby(pd.Grouper(key='Date Installed', freq='Y')).size().reset_index(name='Year Count')
    df2 = pd.DataFrame(year)
    df2['Date Installed'] =df2['Date Installed'].dt.strftime('%Y')

    month = df.groupby(pd.Grouper(key='Date Installed', freq='M')).size().reset_index(name='Month Count')
    df3 = pd.DataFrame(month)
    df3['Date Installed'] =df3['Date Installed'].dt.strftime('%Y-%m')
    df3 = df3.sort_values(by=['Date Installed'], ascending=True)

    df['Date Installed'] = df['Date Installed'].dt.strftime('%Y-%m')
    bor = df.groupby(['Date Installed', 'Borough']).size().reset_index(name='Borough Count')
    df4 = pd.DataFrame(bor)

    result = pd.concat([df2, df3, df4], ignore_index=True)
    result = result.sort_values(by=['Date Installed', 'Month Count'], ascending=[True, True])
    result['dup'] = result['Date Installed'].shift(1)
    result['Date Installed'] = result.apply(lambda x: np.nan if x['Date Installed'] == x['dup'] else x['Date Installed'], axis=1)
    result = result.drop('dup', axis=1)
    return result

def main():
    brooklyn = table1('Brooklyn')
    manhattan = table1('Manhattan')
    queens = table1('Queens')
    statenisland = table1('Staten Island')
    bronx = table1('the Bronx')
    frames = [brooklyn, manhattan, queens, statenisland, bronx]

    with pd.ExcelWriter('dataintables.xlsx') as writer:

        t1result = pd.concat(frames)
        t1result.to_excel(writer,index=False, sheet_name='Table 1')

        t2result = table2()
        t2result.to_excel(writer,index=False, sheet_name='Table 2')

main()
