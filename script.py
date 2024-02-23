import pandas as pd  
import numpy as np

def wczytanie(x):
    sciezka_do_pliku = 'Plik_do_zadania.xlsx'

    try:
        df = pd.read_excel(sciezka_do_pliku, sheet_name=x)
        print(df)
    except Exception as e:
        print(f"Wystąpił błąd podczas wczytywania pliku: {e}")
    return df

def konwersja(df, podmiana_df):
    df = df.drop(0)

    nowy_df = pd.DataFrame(columns=['tourists_type', 'period_value', 'period', 'N'])

    for index, row in df.iterrows():
        nazwa = df.iloc[2:11, 0].tolist()
        nazwa = podmiana(nazwa, podmiana_df)
        okresy = df.iloc[1, 1:10].tolist()
        kategorie = df.iloc[0, 1:10].tolist()
        kategorie = [value for value in kategorie if not isinstance(value, float) or not np.isnan(value)]
        wartosci = df.iloc[2:11, 1:10]

    
    for j in range(len(nazwa)):
        i = 0 
        wartosci_dla = wartosci.iloc[j, 0:9].tolist()
        for okres in okresy:
            if i < 6:
                kategoria = kategorie[0] 
            elif i < 8:
                kategoria = kategorie[1]
            else:
                kategoria = kategorie[2]
            nowy_wiersz = {'tourists_type': nazwa[j],'period_value': kategoria, 'period': okres, 'N': wartosci_dla[i]}
            nowy_df = pd.concat([nowy_df, pd.DataFrame([nowy_wiersz])], ignore_index=True)
            i+=1





    nowa_nazwa_pliku = 'nowa_nazwa_pliku.xlsx'
    nowy_df.to_excel(nowa_nazwa_pliku, index=False)

    print(f'Plik {nowa_nazwa_pliku} został utworzony pomyślnie.')

def podmiana(nazwa, podmiana_df):
    translation_dict = dict(zip(podmiana_df['ANG'], podmiana_df['PL']))

    translated_list = [translation_dict.get(value, value) for value in nazwa]

    return translated_list


def main():
    dane = wczytanie('Number of tourist')
    podmiana_df = wczytanie('Dictionary')
    konwersja(dane, podmiana_df)

if __name__ == "__main__":
    main()
