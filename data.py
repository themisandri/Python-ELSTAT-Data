import xlrd
from matplotlib import pyplot as plt
import mysql.connector
import pandas as pd
from sqlalchemy import create_engine
import pymysql
import seaborn as sns


book = xlrd.open_workbook("excel/Αφίξεις ανά μέσο 4ο 2011.xls")
book2 = xlrd.open_workbook("excel/Αφίξεις ανά μέσο 4ο 2012.xls")
book3 = xlrd.open_workbook("excel/Αφίξεις ανά μέσο 4ο 2013.xls")
book4 = xlrd.open_workbook("excel/Αφίξεις ανά μέσο 4ο 2014.xls")
book5 = xlrd.open_workbook("excel/Αφίξεις ανά μέσο 4ο 2015.xls")

def get_tourists(save_to_db=False):
    #Save the sheets to dataframes
    sheet = book.sheet_by_index(11)
    total2011 = sheet.cell(134, 6).value

    sheet2 = book2.sheet_by_index(11)
    total2012 = sheet2.cell(136, 6).value

    sheet3 = book3.sheet_by_index(11)
    total2013 = sheet3.cell(136, 6).value

    sheet4 = book4.sheet_by_index(11)
    total2014 = sheet4.cell(136, 6).value

    sheet5 = book5.sheet_by_index(11)
    total2015 = sheet5.cell(136, 6).value

    #create dataframe with total tourists per year
    df = pd.DataFrame({'Year': [2011, 2012, 2013, 2014, 2015], 'Total': [total2011, total2012, total2013, total2014, total2015]})
    #year column to datetime
    df['Year'] = pd.to_datetime(df['Year'], format='%Y')

    #plot the data in line chart and prevent scientific notation
    fig = plt.figure()
    ax = fig.add_subplot(111)
    ax.plot(df['Year'], df['Total'])
    ax.yaxis.get_major_formatter().set_scientific(False)
    plt.xticks(rotation=45)
    plt.title('Total tourists per year')
    plt.xlabel('Year')
    plt.ylabel('Total tourists')
    fig = plt.gcf()


    if save_to_db == True:
        #Store to database
        database = mysql.connector.connect(host="localhost", user="root", passwd="", db="mysqlPython")
        cursor = database.cursor()
        query = 'INSERT INTO total (tourists,year) VALUES (%s,%s)'
        records_to_insert = [(total2011, 2011),
                            (total2012, 2012),
                            (total2013, 2013),
                            (total2014, 2014),
                            (total2015, 2015)]
        cursor.executemany(query, records_to_insert)
        cursor.close()

        # Commit the transaction
        database.commit()
        database.close()

    return fig

def per_country(save_to_db=False):
    x2011 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2011.xls")
    x2012 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2012.xls")
    x2013 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2013.xls")
    x2014 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2014.xls")
    x2015 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2015.xls")

    cols = [1, 6]
    df2011 = pd.read_excel(x2011, "ΔΕΚ", usecols=cols)

    df2011.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car", "Unnamed: 6": "total"}, inplace=True)

    var = df2011.iloc[78:132, 1:6]
    max = var['total'].max()
    xwra2011 = df2011.loc[df2011['total'] == max]
    val2011 = xwra2011.iloc[:, 0]
    total2011 = xwra2011.iloc[:, 1]

    df2012 = pd.read_excel(x2012, "ΔΕΚ", usecols=cols)

    df2012.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car", "Unnamed: 6": "total"}, inplace=True)

    var = df2012.iloc[80:134, 1:6]
    max = var['total'].max()
    xwra2012 = df2012.loc[df2012['total'] == max]
    val2012 = xwra2012.iloc[:, 0]
    total2012 = xwra2012.iloc[:, 1]

    df2013 = pd.read_excel(x2013, "ΔΕΚ", usecols=cols)

    df2013.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car", "Unnamed: 6": "total"}, inplace=True)

    var = df2013.iloc[80:134, 1:6]
    max = var['total'].max()
    xwra2013 = df2013.loc[df2013['total'] == max]
    val2013 = xwra2013.iloc[:, 0]
    total2013 = xwra2013.iloc[:, 1]

    df2014 = pd.read_excel(x2014, "ΔΕΚ", usecols=cols)

    df2014.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car", "Unnamed: 6": "total"}, inplace=True)

    var = df2014.iloc[80:134, 1:6]
    max = var['total'].max()
    xwra2014 = df2014.loc[df2014['total'] == max]
    val2014 = xwra2014.iloc[:, 0]
    total2014 = xwra2014.iloc[:, 1]

    df2015 = pd.read_excel(x2015, "ΔΕΚΕΜ", usecols=cols)

    df2015.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car", "Unnamed: 6": "total"}, inplace=True)

    var = df2015.iloc[80:134, 1:6]
    max = var['total'].max()
    xwra2015 = df2015.loc[df2015['total'] == max]
    val2015 = xwra2015.iloc[:, 0]
    total2015 = xwra2015.iloc[:, 1]

    #add year column to dataframes
    xwra2011['year'] = 2011
    xwra2012['year'] = 2012
    xwra2013['year'] = 2013
    xwra2014['year'] = 2014
    xwra2015['year'] = 2015

    if save_to_db == True:
        # create sqlalchemy engine
        engine = create_engine("mysql+pymysql://{user}:{pw}@localhost/{db}"
                            .format(user="root",
                                    pw="",
                                    db="mysqlPython"))

        xwra2011.to_sql('countries', con=engine, if_exists='append', chunksize=1000)
        xwra2012.to_sql('countries', con=engine, if_exists='append', chunksize=1000)
        xwra2013.to_sql('countries', con=engine, if_exists='append', chunksize=1000)
        xwra2014.to_sql('countries', con=engine, if_exists='append', chunksize=1000)
        xwra2015.to_sql('countries', con=engine, if_exists='append', chunksize=1000)

    frames = [xwra2011, xwra2012, xwra2013, xwra2014, xwra2015]
    result = pd.concat(frames)
    result = result.reset_index(drop=True)
    result['year'] = pd.to_datetime(result['year'], format='%Y')

    #plot with seaborn
    fig = plt.figure()
    sns.set(style="whitegrid")
    sns.set(rc={'figure.figsize': (11.7, 8.27)})
    ax = sns.barplot(x="year", y="total", hue="countries", data=result)
    plt.xticks(rotation=45)
    ax.get_yaxis().get_major_formatter().set_scientific(False)
    plt.title('Countries of origin with the largest share of tourist arrivals')
    plt.xlabel('Year')
    plt.ylabel('Total tourists')
    plt.show()

    return fig

def per_month(save_to_db=False):
    jan11 = book.sheet_by_index(0)
    feb11 = book.sheet_by_index(1)
    mar11 = book.sheet_by_index(2)
    apr11 = book.sheet_by_index(3)
    may11 = book.sheet_by_index(4)
    jun11 = book.sheet_by_index(5)
    jul11 = book.sheet_by_index(6)
    aug11 = book.sheet_by_index(7)
    sep11 = book.sheet_by_index(8)
    oct11 = book.sheet_by_index(9)
    nov11 = book.sheet_by_index(10)
    dec11 = book.sheet_by_index(11)

    jan12 = book2.sheet_by_index(0)
    feb12 = book2.sheet_by_index(1)
    mar12 = book2.sheet_by_index(2)
    apr12 = book2.sheet_by_index(3)
    may12 = book2.sheet_by_index(4)
    jun12 = book2.sheet_by_index(5)
    jul12 = book2.sheet_by_index(6)
    aug12 = book2.sheet_by_index(7)
    sep12 = book2.sheet_by_index(8)
    oct12 = book2.sheet_by_index(9)
    nov12 = book2.sheet_by_index(10)
    dec12 = book2.sheet_by_index(11)

    jan13 = book3.sheet_by_index(0)
    feb13 = book3.sheet_by_index(1)
    mar13 = book3.sheet_by_index(2)
    apr13 = book3.sheet_by_index(3)
    may13 = book3.sheet_by_index(4)
    jun13 = book3.sheet_by_index(5)
    jul13 = book3.sheet_by_index(6)
    aug13 = book3.sheet_by_index(7)
    sep13 = book3.sheet_by_index(8)
    oct13 = book3.sheet_by_index(9)
    nov13 = book3.sheet_by_index(10)
    dec13 = book3.sheet_by_index(11)

    jan14 = book4.sheet_by_index(0)
    feb14 = book4.sheet_by_index(1)
    mar14 = book4.sheet_by_index(2)
    apr14 = book4.sheet_by_index(3)
    may14 = book4.sheet_by_index(4)
    jun14 = book4.sheet_by_index(5)
    jul14 = book4.sheet_by_index(6)
    aug14 = book4.sheet_by_index(7)
    sep14 = book4.sheet_by_index(8)
    oct14 = book4.sheet_by_index(9)
    nov14 = book4.sheet_by_index(10)
    dec14 = book4.sheet_by_index(11)

    jan15 = book5.sheet_by_index(0)
    feb15 = book5.sheet_by_index(1)
    mar15 = book5.sheet_by_index(2)
    apr15 = book5.sheet_by_index(3)
    may15 = book5.sheet_by_index(4)
    jun15 = book5.sheet_by_index(5)
    jul15 = book5.sheet_by_index(6)
    aug15 = book5.sheet_by_index(7)
    sep15 = book5.sheet_by_index(8)
    oct15 = book5.sheet_by_index(9)
    nov15 = book5.sheet_by_index(10)
    dec15 = book5.sheet_by_index(11)



    months11 = [(jan11.cell(65, 6).value, jan11.cell(1, 1).value), (feb11.cell(65, 6).value, feb11.cell(1, 1).value), (mar11.cell(65, 6).value, mar11.cell(1, 1).value), (apr11.cell(65, 6).value, apr11.cell(1, 1).value),
                (may11.cell(65, 6).value, may11.cell(1, 1).value), (jun11.cell(65, 6).value, jun11.cell(1, 1).value), (jul11.cell(65, 6).value, jul11.cell(1, 1).value), (aug11.cell(65, 6).value, aug11.cell(1, 1).value),
                (sep11.cell(65, 6).value, sep11.cell(1, 1).value), (oct11.cell(65, 6).value, oct11.cell(1, 1).value), (nov11.cell(65, 6).value, nov11.cell(1, 1).value), (dec11.cell(65, 6).value, dec11.cell(1, 1).value)]

    months12 = [(jan12.cell(65, 6).value, jan12.cell(1, 1).value), (feb12.cell(65, 6).value, feb12.cell(1, 1).value), (mar12.cell(65, 6).value, mar12.cell(1, 1).value), (apr12.cell(65, 6).value, apr12.cell(1, 1).value),
                (may12.cell(65, 6).value, may12.cell(1, 1).value), (jun12.cell(65, 6).value, jun12.cell(1, 1).value), (jul12.cell(65, 6).value, jul12.cell(1, 1).value), (aug12.cell(65, 6).value, aug12.cell(1, 1).value),
                (sep12.cell(65, 6).value, sep12.cell(1, 1).value), (oct12.cell(65, 6).value, oct12.cell(1, 1).value), (nov12.cell(65, 6).value, nov12.cell(1, 1).value), (dec12.cell(65, 6).value, dec12.cell(1, 1).value)]

    months13 = [(jan13.cell(64, 6).value, jan13.cell(1, 1).value), (feb13.cell(64, 6).value, feb13.cell(1, 1).value), (mar13.cell(64, 6).value, mar13.cell(1, 1).value), (apr13.cell(64, 6).value, apr13.cell(1, 1).value),
                (may13.cell(64, 6).value, may13.cell(1, 1).value), (jun13.cell(64, 6).value, jun13.cell(1, 1).value), (jul13.cell(65, 6).value, jul13.cell(1, 1).value), (aug13.cell(65, 6).value, aug13.cell(1, 1).value),
                (sep13.cell(65, 6).value, sep13.cell(1, 1).value), (oct13.cell(65, 6).value, oct13.cell(1, 1).value), (nov13.cell(65, 6).value, nov13.cell(1, 1).value), (dec13.cell(65, 6).value, dec13.cell(1, 1).value)]

    months14 = [(jan14.cell(65, 6).value, jan14.cell(1, 1).value), (feb14.cell(65, 6).value, feb14.cell(1, 1).value), (mar14.cell(65, 6).value, mar14.cell(1, 1).value), (apr14.cell(65, 6).value, apr14.cell(1, 1).value),
                (may14.cell(65, 6).value, may14.cell(1, 1).value), (jun14.cell(65, 6).value, jun14.cell(1, 1).value), (jul14.cell(65, 6).value, jul14.cell(1, 1).value), (aug14.cell(65, 6).value, aug14.cell(1, 1).value),
                (sep14.cell(65, 6).value, sep14.cell(1, 1).value), (oct14.cell(65, 6).value, oct14.cell(1, 1).value), (nov14.cell(65, 6).value, nov14.cell(1, 1).value), (dec14.cell(65, 6).value, dec14.cell(1, 1).value)]

    months15 = [(jan15.cell(66, 6).value, jan15.cell(1, 1).value), (feb15.cell(66, 6).value, feb15.cell(1, 1).value), (mar15.cell(66, 6).value, mar15.cell(1, 1).value), (apr15.cell(66, 6).value, apr15.cell(1, 1).value),
                (may15.cell(66, 6).value, may15.cell(1, 1).value), (jun15.cell(66, 6).value, jun15.cell(1, 1).value), (jul15.cell(66, 6).value, jul15.cell(1, 1).value), (aug15.cell(66, 6).value, aug15.cell(1, 1).value),
                (sep15.cell(66, 6).value, sep15.cell(1, 1).value), (oct15.cell(66, 6).value, oct15.cell(1, 1).value), (nov15.cell(66, 6).value, nov15.cell(1, 1).value), (dec15.cell(66, 6).value, dec15.cell(1, 1).value)]

    if save_to_db == True:

        database = mysql.connector.connect(host="localhost", user="root", passwd="", db="mysqlPython")
        cursor = database.cursor()

        query = 'INSERT INTO months (total,month) VALUES (%s,%s)'
        values = months11
        cursor.executemany(query, values)
        query = 'INSERT INTO months (total,month) VALUES (%s,%s)'
        values = months12
        cursor.executemany(query, values)
        query = 'INSERT INTO months (total,month) VALUES (%s,%s)'
        values = months13
        cursor.executemany(query, values)
        query = 'INSERT INTO months (total,month) VALUES (%s,%s)'
        values = months14
        cursor.executemany(query, values)
        query = 'INSERT INTO months (total,month) VALUES (%s,%s)'
        values = months15
        cursor.executemany(query, values)
        cursor.close()

        # Commit the transaction
        database.commit()
        database.close()

    months = months11 + months12 + months13 + months14 + months15
    # plot
    #plt.ticklabel_format(style='plain')

    xs = [x[1] for x in months]
    ys = [x[0] for x in months]
    plt.xlabel('Month')
    plt.xticks(rotation=90)
    plt.ylabel('Number of tourists')
    plt.title('Tourists per Month')
    #size increase
    plt.rcParams["figure.figsize"] = [16,9]
    #gap between bars
    plt.rcParams['xtick.major.pad']='8'
    plt.plot(xs, ys, 'ro')
    plt.plot(xs, ys)
    fig = plt.gcf()
    plt.show()

    return fig

def per_vehicle(save_to_db=False):
    x2011 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2011.xls")
    x2012 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2012.xls")
    x2013 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2013.xls")
    x2014 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2014.xls")
    x2015 = pd.ExcelFile("excel/Αφίξεις ανά μέσο 4ο 2015.xls")

    cols = [1, 2, 3, 4, 5, 6]
    df2011 = pd.read_excel(x2011, "ΔΕΚ", usecols=cols)
    vehicle2011 = df2011.iloc[[133], 1:5]
    vehicle2011['year'] = 2011
    vehicle2011.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car"}, inplace=True)

    df2012 = pd.read_excel(x2012, "ΔΕΚ", usecols=cols)
    vehicle2012 = df2012.iloc[[135], 1:5]
    vehicle2012['year'] = 2012
    vehicle2012.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car"}, inplace=True)

    df2013 = pd.read_excel(x2013, "ΔΕΚ", usecols=cols)
    vehicle2013 = df2013.iloc[[135], 1:5]
    vehicle2013['year'] = 2013
    vehicle2013.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car"}, inplace=True)

    df2014 = pd.read_excel(x2014, "ΔΕΚ", usecols=cols)
    vehicle2014 = df2014.iloc[[135], 1:5]
    vehicle2014['year'] = 2014
    vehicle2014.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car"}, inplace=True)

    df2015 = pd.read_excel(x2015, "ΔΕΚΕΜ", usecols=cols)
    vehicle2015 = df2015.iloc[[135], 1:5]
    vehicle2015['year'] = 2015
    vehicle2015.rename(
        columns={"Unnamed: 1": "countries", "Unnamed: 2": "plane", "Unnamed: 3": "train", "Unnamed: 4": "boat",
                 "Unnamed: 5": "car"}, inplace=True)

    vehicle2011['year'] = pd.to_datetime(vehicle2011['year'], format='%Y')
    vehicle2012['year'] = pd.to_datetime(vehicle2012['year'], format='%Y')
    vehicle2013['year'] = pd.to_datetime(vehicle2013['year'], format='%Y')
    vehicle2014['year'] = pd.to_datetime(vehicle2014['year'], format='%Y')
    vehicle2015['year'] = pd.to_datetime(vehicle2015['year'], format='%Y')

    if save_to_db == True:

        # create sqlalchemy engine
        engine = create_engine("mysql+pymysql://{user}:{pw}@localhost/{db}"
                            .format(user="root",
                                    pw="",
                                    db="mysqlPython"))

        vehicle2011.to_sql('vehicle', con=engine, if_exists='append', chunksize=1000, )
        vehicle2012.to_sql('vehicle', con=engine, if_exists='append', chunksize=1000)
        vehicle2013.to_sql('vehicle', con=engine, if_exists='append', chunksize=1000)
        vehicle2014.to_sql('vehicle', con=engine, if_exists='append', chunksize=1000)
        vehicle2015.to_sql('vehicle', con=engine, if_exists='append', chunksize=1000)

    # plot
    frames = [vehicle2011, vehicle2012, vehicle2013, vehicle2014, vehicle2015]
    result = pd.concat(frames)

    result.plot(x='year', y='plane', color='red')
    fig1 = plt.gcf()
    result.plot(x='year', y='train', color='red')
    fig2 = plt.gcf()
    result.plot(x='year', y='boat', color='red')
    fig3 = plt.gcf()
    result.plot(x='year', y='car', color='red')
    fig4 = plt.gcf()
    plt.show()

    return fig1, fig2, fig3, fig4
