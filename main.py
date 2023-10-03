import pandas as pd
import requests


organisasjons_nummer = [
    953376040,
    960010833,
    986351620,
    961702615,
    991434682,
    912304752,
    997864417,
    916231539,
    986018522,
    937920040,
    957345808
]


def hent_enhet(organisasjonsnummer):
    url = f"https://data.brreg.no/enhetsregisteret/api/enheter/{organisasjonsnummer}"
    try:
        response = requests.get(url)

        if response.status_code == 200:
            data = response.json()
            return data
        else:
            return None
    except requests.exceptions.RequestException as e:
        print(e)
        return None
    except requests.exceptions.JSONDecodeError as e:
        print(e)
        return None


def hent_regnskap(organisasjonsnummer):
    url = f"https://data.brreg.no/regnskapsregisteret/regnskap/{organisasjonsnummer}"
    try:
        response = requests.get(url)

        if response.status_code == 200:
            data = response.json()
            return data
        else:
            return None
    except requests.exceptions.RequestException as e:
        print(e)
        return None
    except requests.exceptions.JSONDecodeError as e:
        print(e)
        return None


def hent_alle_enheter():
    alle_enheter = []
    for org_nummer in organisasjons_nummer:
        enhet = hent_enhet(org_nummer)
        alle_enheter.append(enhet)
    return alle_enheter


def hent_alle_regnskap():
    alle_regnskap = []
    for org_nummer in organisasjons_nummer:
        regnskap = hent_regnskap(org_nummer)
        alle_regnskap.append(regnskap)
    return alle_regnskap


def lag_dataframe(liste):
    liste_med_df = []
    for element in liste:
        df = pd.DataFrame(element)
        liste_med_df.append(df)
    return liste_med_df


def clean_dataframe(liste):
    for df in liste:
        df = df.dropna()
        df = df.drop_duplicates()


def skriv_til_excel(liste_med_df, liste_med_df2):
    with pd.ExcelWriter('output.xlsx') as writer:
        for i, df in enumerate(liste_med_df):
            df.to_excel(writer, sheet_name=f'Regnskap {i+1}')
        for i, df in enumerate(liste_med_df2):
            df.to_excel(writer, sheet_name=f'Enheter {i+1}')


def main():
    alle_regnskap = hent_alle_regnskap()
    alle_enheter = hent_alle_enheter()
    liste_med_regnskapdf = lag_dataframe(alle_regnskap)
    liste_med_enhetdf = lag_dataframe(alle_enheter)
    clean_dataframe(liste_med_enhetdf)

    skriv_til_excel(liste_med_regnskapdf, liste_med_enhetdf)


main()
