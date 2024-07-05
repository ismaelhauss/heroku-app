from django.shortcuts import render
from BP_generator.forms import BP_form
from BP_generator.utils import get_new_excel, get_chart, make_important

from bs4 import BeautifulSoup

def data_form(request):
    return render(request, "BP_generator/form.html")


def contact(request):
    displayed_page = 'BP_generator/contact.html'
    if request.method == 'POST':
        form = BP_form(request.POST)  # ajout d’un nouveau formulaire ici
        if form.is_valid():
            print("VALID")
            tab = []
            chart = []
            label_prefix = "EBIT "
            form = get_new_excel("BP Marché touristique.xlsx", "output.xlsx", "Tableau résultat ", request.POST)
            x = form.columns[1:]
            ys = form.iloc(0)
            form_pessimiste = get_new_excel("BP Marché touristique.xlsx", "output.xlsx", "Tableau résultat pessimiste", request.POST)
            ys_pessimiste = form_pessimiste.iloc(0)
            form_raisonnable = get_new_excel("BP Marché touristique.xlsx", "output.xlsx", "Tableau résultat raisonnable", request.POST)
            ys_raisonnable = form_raisonnable.iloc(0)
            form_optimiste = get_new_excel("BP Marché touristique.xlsx", "output.xlsx", "Tableau résultat optimiste", request.POST)
            ys_optimiste = form_optimiste.iloc(0)

            # Création du graphique
            #chart.append(get_chart('Evolution de l\'Ebida', x, ys[7][1:]))
            chart.append(get_chart(
                'Evolution de l\'EBIT', 
                x, ys[9][1:], 
                label_prefix+"désiré", 
                ys_pessimiste[9][1:], 
                ys_raisonnable[9][1:], 
                ys_optimiste[9][1:], 
                names=(label_prefix+"pessimiste", label_prefix+"raisonnable", label_prefix+"optimiste",))
            )
            
            form = form.to_html()
            form = BeautifulSoup(form, 'html.parser')

            # Ajouter un style à la deuxième ligne (index 1)
            rows = form.find_all('tr')
            rows[1]['style'] = 'background-color: #9f9;'
            rows[3]['style'] = 'background-color: #f99;'
            rows[8]['style'] = 'background-color: #77f;'
            rows[10]['style'] = 'background-color: #7f7;'

            # Convertir le HTML modifié en chaîne
            form = str(form)
            tab.append(form)
            displayed_page = 'BP_generator/bp.html'
            
            return render(request,
                displayed_page,
                {'tab': tab, 'chart': chart},)  # passe ce formulaire au gabarit
    else:   
        form = BP_form()  # ajout d’un nouveau formulaire ici
    return render(request,
        displayed_page,
        {'form': form})  # passe ce formulaire au gabarit