# 🐍 xlwings — Python dans Excel, sans VBA

VBA existe depuis 1993. Il fait le travail — mais il est verbeux, difficile à déboguer,
impossible à tester unitairement, et coupé de tout l'écosystème Python moderne.

Ce repo documente mon utilisation de **xlwings** pour remplacer des macros VBA
par du Python propre, et exposer des fonctions Python directement dans les cellules Excel
comme si c'étaient des fonctions natives.

---

## Ce que xlwings permet concrètement

```
  ┌──────────────────────────────────────────────────────────┐
  │                    Fichier Excel                         │
  │                                                          │
  │  Cellule A1: =square(B1)    ← fonction Python UDF       │
  │  Cellule A2: =forecast(C1:C12)  ← pandas sous le capot  │
  │                                                          │
  │  [Bouton "Générer Rapport"]  ← déclenche script Python  │
  │                                                          │
  └──────────────────────────────────────────────────────────┘
         │                              │
         │ UDF call                     │ Macro call
         ▼                              ▼
  ┌─────────────────────┐    ┌─────────────────────────────┐
  │   functions.py      │    │      report.py              │
  │                     │    │                             │
  │  @xw.func           │    │  def generate_report():     │
  │  def square(x):     │    │    wb = xw.Book.caller()    │
  │      return x ** 2  │    │    sheet = wb.sheets[0]     │
  │                     │    │    df = get_data()          │
  │  @xw.func           │    │    sheet.range('A1').value  │
  │  @xw.arg('data',    │    │      = df                  │
  │    pd.DataFrame)    │    │    add_chart(sheet, df)     │
  │  def forecast(data):│    │                             │
  │    return model(data│    └─────────────────────────────┘
  │                     │
  └─────────────────────┘
```

---

## Cas d'usage 1 : UDF Python dans Excel

Écrire une fonction Python et l'appeler depuis une cellule Excel comme `=SOMME()`.

```python
import xlwings as xw
import pandas as pd
import numpy as np

@xw.func
def square(x):
    """Retourne le carré de x — utilisable en cellule: =square(A1)"""
    return x ** 2

@xw.func
@xw.arg('data', pd.DataFrame, index=False, header=True)
@xw.ret(index=False, header=True)
def moving_average(data, window: int = 3):
    """Moyenne mobile sur un plage Excel — =moving_average(A1:A20, 5)"""
    return data.rolling(window=window).mean()

@xw.func
@xw.arg('prices', np.array)
def volatility(prices):
    """Volatilité annualisée — =volatility(B2:B252)"""
    returns = np.diff(np.log(prices))
    return np.std(returns) * np.sqrt(252)
```

---

## Cas d'usage 2 : Automatisation complète d'un rapport

Bouton Excel → script Python génère tout le rapport en quelques secondes.

```python
import xlwings as xw
import pandas as pd
import matplotlib.pyplot as plt

def generate_monthly_report():
    wb = xw.Book.caller()  # classeur qui a déclenché la macro
    ws_data  = wb.sheets['Data']
    ws_report = wb.sheets['Rapport']

    # 1. Lire les données depuis la feuille Data
    df = ws_data.range('A1').expand().options(pd.DataFrame).value

    # 2. Calculer les KPIs
    summary = df.groupby('Catégorie').agg(
        Total=('Montant', 'sum'),
        Moyenne=('Montant', 'mean'),
        Count=('ID', 'count')
    ).reset_index()

    # 3. Écrire les résultats dans la feuille Rapport
    ws_report.range('B2').value = summary

    # 4. Créer un graphique matplotlib et l'insérer dans Excel
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar(summary['Catégorie'], summary['Total'], color='steelblue')
    ax.set_title('Ventes par catégorie')
    ws_report.pictures.add(fig, name='SalesChart',
                           update=True,
                           left=ws_report.range('B15').left,
                           top=ws_report.range('B15').top)

    print("Rapport généré avec succès.")
```

---

## Installation & setup

```bash
# 1. Installer xlwings
pip install xlwings

# 2. Installer l'add-in Excel
xlwings addin install

# 3. Démarrer un projet
xlwings quickstart mon_projet
# → crée mon_projet.xlsb + mon_projet.py liés automatiquement

# 4. Importer les fonctions UDF
# Dans Excel : onglet xlwings → Import Functions
```

---

## Pourquoi Python > VBA pour ce cas

```
  VBA                          Python (xlwings)
  ──────────────────           ─────────────────────────
  Syntaxe des années 90        Syntaxe moderne et lisible
  Pas de pandas / numpy        Accès à tout l'écosystème
  Débogage dans l'IDE Excel    Débogage dans VS Code / PyCharm
  Pas de tests unitaires       pytest sur toutes les fonctions
  Pas de versioning propre     Git fonctionne normalement
  Impossible à dockeriser      Scriptable en CI/CD
```

---

## Ce que j'ai vraiment appris

La subtilité la plus importante : xlwings utilise **COM** sur Windows (`pywin32`)
pour communiquer avec Excel. Ça veut dire que Python et Excel tournent dans deux
processus séparés — Excel appelle Python, attend la réponse, et affiche le résultat.

En pratique : si une UDF plante en Python, Excel affiche `#VALUE!` sans détail.
Le reflex est d'ajouter `try/except` dans chaque UDF et de logger l'erreur dans
un fichier `.log` — sinon le débogage devient un cauchemar.

---

*Projet réalisé dans le cadre de ma formation ingénieur — ENSET Mohammedia*
*Par **Abderrahmane Elouafi** · [LinkedIn](https://www.linkedin.com/in/abderrahmane-elouafi-43226736b/) · [Portfolio](https://my-first-porfolio-six.vercel.app/)*
