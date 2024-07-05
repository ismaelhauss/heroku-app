# listings/forms.py

from django import forms
from datetime import datetime

class BP_form(forms.Form):
   begin_date = forms.DateField(initial=datetime.now().__format__("%d/%m/%Y"))
   nb_velo = forms.FloatField(label="Combien de vélos", initial=5)
   prix_exp = forms.FloatField(label="Prix d'une expérience", initial=40)
   dist_moy_exp = forms.FloatField(label="Distance moyenne par expérience", initial=10)
   nb_exp_sem = forms.FloatField(label="Nombre d'expériences en semaine (lundi à vendredi)", initial=10)
   nb_exp_we = forms.FloatField(label="Combien d'expériences le week-end", initial=10)
   cout_chauffeur = forms.FloatField(label="Coût d'un chauffeur à l'heure", initial=15)
   cout_stock = forms.FloatField(label="Coût de stockage par mois", initial=100)