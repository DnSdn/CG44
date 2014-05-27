# Gérer les imports
import os
import csv
import sys
import fileinput

#Dossier dans lequel on travail
dir = "I:/" 

#Constante des fichiers à ouvrir pour SOC15001
metadata1 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14001_fond_aide_jeunes\metadata.xml"'
xls1 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14001_fond_aide_jeunes\SOC14001_xls\fond_aide_jeunes.xls"'
csvCommune1 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14001_fond_aide_jeunes\Donnees_dspl\slice_communes.csv"'
csvDepartement1 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14001_fond_aide_jeunes\Donnees_dspl\slice_departement.csv"'
csvEPCI1 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14001_fond_aide_jeunes\Donnees_dspl\slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15002
metadata2 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14002_contrat_autonomie/metadata.xml"'
xls2 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14002_contrat_autonomie\SOC14002_xls\contrat_autonomie.xls"'
csvCommune2 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14002_contrat_autonomie\Donnees_dspl\slice_communes.csv"'
csvDepartement2 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14002_contrat_autonomie\Donnees_dspl\slice_departement.csv"'
csvEPCI2 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14002_contrat_autonomie\Donnees_dspl\slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15003
metadata3 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14003_aides_permis_jeunes/metadata.xml"'
xls3 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14003_aides_permis_jeunes\SOC14003_xls\aides_permis_jeunes.xls"'
csvCommune3 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14003_aides_permis_jeunes\Donnees_dspl/slice_communes.csv"'
csvDepartement3 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14003_aides_permis_jeunes\Donnees_dspl/slice_departement.csv"'
csvEPCI3 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14003_aides_permis_jeunes\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15004
metadata4 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14004_agrements_adoption/metadata.xml"'
xls4 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14004_agrements_adoption\SOC14004_xls\agrements_adoption.xls"'
csvCommune4 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14004_agrements_adoption\Donnees_dspl/slice_communes.csv"'
csvDepartement4 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14004_agrements_adoption\Donnees_dspl/slice_departement.csv"'
csvEPCI4 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14004_agrements_adoption\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15005
metadata5 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14005_adoptions/metadata.xml"'
xls5 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14005_adoptions\SOC14005_xls\adoptions.xls"'
csvCommune5 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14005_adoptions\Donnees_dspl/slice_communes.csv"'
csvDepartement5 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14005_adoptions\Donnees_dspl/slice_departement.csv"'
csvEPCI5 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14005_adoptions\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15006
metadata6 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14006_TISF/metadata.xml"'
xls6 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14006_TISF\SOC14006_xls\TISF.xls"'
csvCommune6 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14006_TISF\Donnees_dspl/slice_communes.csv"'
csvDepartement6 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14006_TISF\Donnees_dspl/slice_departement.csv"'
csvEPCI6 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14006_TISF\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15007
metadata7 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14007_aides_edu_admin_SED/metadata.xml"'
xls7 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14007_aides_edu_admin_SED\SOC14007_xls\aides_edu_admin_SED.xls"'
csvCommune7 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14007_aides_edu_admin_SED\Donnees_dspl/slice_communes.csv"'
csvDepartement7 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14007_aides_edu_admin_SED\Donnees_dspl/slice_departement.csv"'
csvEPCI7 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14007_aides_edu_admin_SED\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15008
metadata8 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14008_aides_edu_admin_AED/metadata.xml"'
xls8 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14008_aides_edu_admin_AED\SOC14008_xls\aides_edu_admin_AED.xls"'
csvCommune8 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14008_aides_edu_admin_AED\Donnees_dspl/slice_communes.csv"'
csvDepartement8 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14008_aides_edu_admin_AED\Donnees_dspl/slice_departement.csv"'
csvEPCI8 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14008_aides_edu_admin_AED\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15009
metadata9 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14009_aides_edu_judi_AEMO/metadata.xml"'
xls9 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14009_aides_edu_judi_AEMO\SOC14009_xls\aides_edu_judi_AEMO.xls"'
csvCommune9 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14009_aides_edu_judi_AEMO\Donnees_dspl/slice_communes.csv"'
csvDepartement9 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14009_aides_edu_judi_AEMO\Donnees_dspl/slice_departement.csv"'
csvEPCI9 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14009_aides_edu_judi_AEMO\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata10 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14010_aides_edu_jud_AEIMF/metadata.xml"'
xls10 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14010_aides_edu_jud_AEIMF\SOC14010_xls\aides_edu_jud_AEIMF.xls"'
csvCommune10 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14010_aides_edu_jud_AEIMF\Donnees_dspl/slice_communes.csv"'
csvDepartement10 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14010_aides_edu_jud_AEIMF\Donnees_dspl/slice_departement.csv"'
csvEPCI10 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14010_aides_edu_jud_AEIMF\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata11 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14011_mesures_plac_admin/metadata.xml"'
xls11 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14011_mesures_plac_admin\SOC14011_xls\mesures_plac_admin.xls"'
csvCommune11 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14011_mesures_plac_admin\Donnees_dspl/slice_communes.csv"'
csvDepartement11 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14011_mesures_plac_admin\Donnees_dspl/slice_departement.csv"'
csvEPCI11 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14011_mesures_plac_admin\Donnees_dspl/slice_EPCI.csv"'

#Constante des fichiers à ouvrir pour SOC15010
metadata12 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14012_mesures_plac_judi/metadata.xml"'
xls12 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14012_mesures_plac_judi\SOC14012_xls\mesures_plac_judi.xls"'
csvCommune12 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14012_mesures_plac_judi\Donnees_dspl/slice_communes.csv"'
csvDepartement12 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14012_mesures_plac_judi\Donnees_dspl/slice_departement.csv"'
csvEPCI12 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14012_mesures_plac_judi\Donnees_dspl/slice_EPCI.csv"'

#Constante des fichiers à ouvrir pour SOC15010
metadata13 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14013_mesures_accomp/metadata.xml"'
xls13 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14013_mesures_accomp\SOC14013_xls\mesures_accomp.xls"'
csvCommune13 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14013_mesures_accomp\Donnees_dspl/slice_communes.csv"'
csvDepartement13 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14013_mesures_accomp\Donnees_dspl/slice_departement.csv"'
csvEPCI13 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14013_mesures_accomp\Donnees_dspl/slice_EPCI.csv"'

#Constante des fichiers à ouvrir pour SOC15010
metadata14 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14014_veille_enf_danger/metadata.xml"'
xls14 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14014_veille_enf_danger\SOC14014_xls\veille_enf_danger.xls"'
csvCommune14 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14014_veille_enf_danger\Donnees_dspl/slice_communes.csv"'
csvDepartement14 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14014_veille_enf_danger\Donnees_dspl/slice_departement.csv"'
csvEPCI14 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14014_veille_enf_danger\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata15 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14015_consultations_PMI/metadata.xml"'
xls15 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14015_consultations_PMI\SOC14015_xls\consultations_PMI.xls"'
csvCommune15 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14015_consultations_PMI\Donnees_dspl/slice_communes.csv"'
csvDepartement15 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14015_consultations_PMI\Donnees_dspl/slice_departement.csv"'
csvEPCI15 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14015_consultations_PMI\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata16 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14017_places_protect_enf/metadata.xml"'
xls16 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14017_places_protect_enf\SOC14017_xls\places_protect_enf.xls"'
csvCommune16 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14017_places_protect_enf\Donnees_dspl/slice_communes.csv"'
csvDepartement16 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14017_places_protect_enf\Donnees_dspl/slice_departement.csv"'
csvEPCI16 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14017_places_protect_enf\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata17 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14018_liste_etablissements_ASE/metadata.xml"'
xls17 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14018_liste_etablissements_ASE\SOC14018_xls\liste_etablissements_ASE.xls"'
csvCommune17 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14018_liste_etablissements_ASE\Donnees_dspl/slice_communes.csv"'
csvDepartement17 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14018_liste_etablissements_ASE\Donnees_dspl/slice_departement.csv"'
csvEPCI17 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14018_liste_etablissements_ASE\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata18 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14020_nb_assistants_mat/metadata.xml"'
xls18 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14020_nb_assistants_mat\SOC14020_xls\nb_assistants_mat.xls"'
csvCommune18 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14020_nb_assistants_mat\Donnees_dspl/slice_communes.csv"'
csvDepartement18 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14020_nb_assistants_mat\Donnees_dspl/slice_departement.csv"'
csvEPCI18 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14020_nb_assistants_mat\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata19 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14021_places_ass_mat/metadata.xml"'
xls19 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14021_places_ass_mat\SOC14021_xls\places_ass_mat.xls"'
csvCommune19 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14021_places_ass_mat\Donnees_dspl/slice_communes.csv"'
csvDepartement19 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14021_places_ass_mat\Donnees_dspl/slice_departement.csv"'
csvEPCI19 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14021_places_ass_mat\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata20 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14022_nb_assistants_fam/metadata.xml"'
xls20 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14022_nb_assistants_fam\SOC14022_xls\nb_assistants_fam.xls"'
csvCommune20 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14022_nb_assistants_fam\Donnees_dspl/slice_communes.csv"'
csvDepartement20 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14022_nb_assistants_fam\Donnees_dspl/slice_departement.csv"'
csvEPCI20 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14022_nb_assistants_fam\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata21 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14023_places_ass_fam/metadata.xml"'
xls21 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14023_places_ass_fam\SOC14023_xls\Places_ass_fam.xls"'
csvCommune21 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14023_places_ass_fam\Donnees_dspl/slice_communes.csv"'
csvDepartement21 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14023_places_ass_fam\Donnees_dspl/slice_departement.csv"'
csvEPCI21 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14023_places_ass_fam\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata22 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14024_nbstruct_petite_enf/metadata.xml"'
xls22 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14024_nbstruct_petite_enf\SOC14024_xls\nbstruct_petite_enf.xls"'
csvCommune22 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14024_nbstruct_petite_enf\Donnees_dspl/slice_communes.csv"'
csvDepartement22 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14024_nbstruct_petite_enf\Donnees_dspl/slice_departement.csv"'
csvEPCI22 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14024_nbstruct_petite_enf\Donnees_dspl/slice_EPCI.csv"'


#Constante des fichiers à ouvrir pour SOC15010
metadata23 = r'C:/Program Files/Notepad++/notepad++.exe -x "I:\Social\Enfance_ et_ famille\SOC14025_nb_places_petite_enf/metadata.xml"'
xls23 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14025_nb_places_petite_enf\SOC14025_xls\nbplaces_petite_enf.xls"'
csvCommune23 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14025_nb_places_petite_enf\Donnees_dspl/slice_communes.csv"'
csvDepartement23 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14025_nb_places_petite_enf\Donnees_dspl/slice_departement.csv"'
csvEPCI23 = r'C:/Program Files/OpenOffice.org 3/program/scalc.exe -x "I:\Social\Enfance_ et_ famille\SOC14025_nb_places_petite_enf\Donnees_dspl/slice_EPCI.csv"'



metadata1 = App.open(metadata1)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata1)
metadata1.close()

xls1 = App.open(xls1)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M222) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

#Ouverture de sliceCommune.csv

csvCommune1 = App.open(csvCommune1)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-162,77))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#Ouverture du fichier XLS
click(Pattern("f0nd_aide_ie.png").similar(0.80))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI
csvEPCI1 = App.open(csvEPCI1)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type('v', KeyModifier.CTRL)
type('f', KeyModifier.CTRL)
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-1.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#csvEPCI1.close()

#XLS1
click(Pattern("f0nd_aide_ie.png").similar(0.80))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

#Ouverture du fichier slice_departement SOC160001
App.open(csvDepartement1)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-2.png").similar(0.80).targetOffset(-96,38))
#csvDepartement1.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls1.close()
click(Pattern("f0nd_aide_ie.png").similar(0.80))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))

#SOC 16002

metadata2 = App.open(metadata2)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)              

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata2)
metadata2.close()

xls2 = App.open(xls2)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune2 = App.open(csvCommune2)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#XLS2
click(Pattern("1394012446091.png").similar(0.78))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI
csvEPCI2 = App.open(csvEPCI2)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#XLS2
click(Pattern("1394012446091.png").similar(0.78))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

#Ouverture du fichier slice_departement
App.open(csvDepartement2)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-3.png").similar(0.80).targetOffset(-96,38))
#csvDepartement2.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls2.close()
click(Pattern("1394012446091.png").similar(0.78))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))

#SOC 16003

metadata3 = App.open(metadata3)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        date = input ("Entrer une date de mise a jour (01/01/2013) :")
        print(date)
        paste(date)  

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata3)
metadata3.close()

xls3 = App.open(xls3)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune3 = App.open(csvCommune3)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#Ouverture du fichier XLS
click(Pattern("aidesjermisj.png").similar(0.80))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI
csvEPCI3 = App.open(csvEPCI3)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#Ouverture du fichier XLS
click(Pattern("aidesjermisj.png").similar(0.80))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))



#Ouverture du fichier slice_departement
App.open(csvDepartement3)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-4.png").similar(0.80).targetOffset(-96,38))
#csvDepartement3.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls3.close()
click(Pattern("aidesjermisj.png").similar(0.80))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))

#SOC 16004

metadata4 = App.open(metadata4)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata4)
metadata4.close()

xls4 = App.open(xls4)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune4 = App.open(csvCommune4)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#Ouverture du fichier XLS
click(Pattern("agrements_ad.png").similar(0.84))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI
csvEPCI4 = App.open(csvEPCI4)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#Ouverture du fichier XLS
click(Pattern("agrements_ad.png").similar(0.84))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

#Ouverture du fichier slice_departement
App.open(csvDepartement4)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-5.png").similar(0.80).targetOffset(-96,38))
#csvDepartement4.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls4.close()
click(Pattern("agrements_ad.png").similar(0.84))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#SOC 16005

metadata5 = App.open(metadata5)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata5)
metadata5.close()

xls5 = App.open(xls5)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune5 = App.open(csvCommune5)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#Ouverture du fichier XLS
click(Pattern("ad0ptinsxs0p.png").similar(0.82))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI
csvEPCI5 = App.open(csvEPCI5)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#Ouverture du fichier XLS
click(Pattern("ad0ptinsxs0p.png").similar(0.82))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))



#Ouverture du fichier slice_departement
App.open(csvDepartement5)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-6.png").similar(0.80).targetOffset(-96,38))
#csvDepartement5.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls5.close()
click(Pattern("ad0ptinsxs0p.png").similar(0.82))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#SOC 16006
metadata6 = App.open(metadata6)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata6)
metadata6.close()

xls6 = App.open(xls6)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune6 = App.open(csvCommune6)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#XLS6
click(Pattern("TlSFxsOpan0f.png").similar(0.83))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 16006
csvEPCI6 = App.open(csvEPCI6)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#XLS6
click(Pattern("TlSFxsOpan0f.png").similar(0.83))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 16006
App.open(csvDepartement6)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement6.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls6.clclick()
click(Pattern("TlSFxsOpan0f.png").similar(0.83))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))



#SOC 14007
metadata7 = App.open(metadata7)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata7)
metadata7.close()

xls7 = App.open(xls7)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune7 = App.open(csvCommune7)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#XLS7
click(Pattern("ISIEJEEIEISI.png").similar(0.94))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15008
csvEPCI7 = App.open(csvEPCI7)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#XLS7
click(Pattern("ISIEJEEIEISI.png").similar(0.94))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15008
App.open(csvDepartement7)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement7.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls7.clclick()
click(Pattern("ISIEJEEIEISI.png").similar(0.94))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))



#SOC 14008
metadata8 = App.open(metadata8)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata8)
metadata8.close()

xls8 = App.open(xls8)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune8 = App.open(csvCommune8)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls8
click(Pattern("ISIEJEEIEISI-1.png").similar(0.87))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15009
csvEPCI8 = App.open(csvEPCI8)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls8
click(Pattern("ISIEJEEIEISI-1.png").similar(0.87))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15009
App.open(csvDepartement8)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement8.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls8.clclick()
click(Pattern("ISIEJEEIEISI-1.png").similar(0.87))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))



#SOC 14009
metadata9 = App.open(metadata9)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata9)
metadata9.close()

xls9 = App.open(xls9)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune9 = App.open(csvCommune9)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls8
click(Pattern("aides_edu_iu.png").similar(0.93))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15010
csvEPCI9 = App.open(csvEPCI9)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls9
click(Pattern("aides_edu_iu.png").similar(0.93))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15010
App.open(csvDepartement9)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement9.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls9.clclick()
click(Pattern("aides_edu_iu.png").similar(0.93))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#SOC 14010
metadata10 = App.open(metadata10)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata10)
metadata10.close()

xls10 = App.open(xls10)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune10 = App.open(csvCommune10)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls10
click(Pattern("aides_edu_iu-1.png").similar(0.88))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15011
csvEPCI10 = App.open(csvEPCI10)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls10
click(Pattern("aides_edu_iu-1.png").similar(0.88))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15011
App.open(csvDepartement10)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement10.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls10.clclick()
click(Pattern("aides_edu_iu-1.png").similar(0.88))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14011
metadata11 = App.open(metadata11)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata11)
metadata11.close()

xls11 = App.open(xls11)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune11 = App.open(csvCommune11)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls11
click(Pattern("mesuresJlac_.png").similar(0.90))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15011
csvEPCI11 = App.open(csvEPCI11)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls11
click(Pattern("mesuresJlac_.png").similar(0.90))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15011
App.open(csvDepartement11)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement11.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls11.clclick()
click(Pattern("mesuresJlac_.png").similar(0.90))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))



#SOC 14012
metadata12 = App.open(metadata12)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata12)
metadata12.close()

xls12 = App.open(xls12)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune12 = App.open(csvCommune12)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls12
click(Pattern("mesuresJJac_.png").similar(0.90))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI12 = App.open(csvEPCI12)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls12
click(Pattern("mesuresJJac_.png").similar(0.90))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement12)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement12.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls12.clclick()
click(Pattern("mesuresJJac_.png").similar(0.90))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14013
metadata13 = App.open(metadata13)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata13)
metadata13.close()

xls13 = App.open(xls13)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune13 = App.open(csvCommune13)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls13
click(Pattern("mesures_acc0.png").similar(0.83))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI13 = App.open(csvEPCI13)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls13
click(Pattern("mesures_acc0.png").similar(0.83))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement13)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement13.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls13.clclick()
click(Pattern("mesures_acc0.png").similar(0.83))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14014
metadata14 = App.open(metadata14)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata14)
metadata14.close()

xls14 = App.open(xls14)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune14 = App.open(csvCommune14)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls14
click(Pattern("veie_enF_dan.png").similar(0.92))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI14 = App.open(csvEPCI14)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls14
click(Pattern("veie_enF_dan.png").similar(0.92))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 14014
App.open(csvDepartement14)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement14.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls14.clclick()
click(Pattern("veie_enF_dan.png").similar(0.92))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14015
metadata15 = App.open(metadata15)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata15)
metadata15.close()

xls15 = App.open(xls15)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune15 = App.open(csvCommune15)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls15
click(Pattern("cnsutatins_P.png").similar(0.85))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI15 = App.open(csvEPCI15)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls15
click(Pattern("cnsutatins_P.png").similar(0.85))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement15)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement15.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls15.clclick()
click(Pattern("cnsutatins_P.png").similar(0.85))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14016
metadata16 = App.open(metadata16)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata16)
metadata16.close()

xls16 = App.open(xls16)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune16 = App.open(csvCommune16)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls16
click(Pattern("pacasJr0tect.png").similar(0.81))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI16 = App.open(csvEPCI16)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls16
click(Pattern("pacasJr0tect.png").similar(0.81))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement16)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement16.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls16.clclick()
click(Pattern("pacasJr0tect.png").similar(0.81))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))



#SOC 14017
metadata17 = App.open(metadata17)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata17)
metadata17.close()

xls17 = App.open(xls17)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune17 = App.open(csvCommune17)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls17
click(Pattern("iste_etabiss.png").similar(0.88))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI17 = App.open(csvEPCI17)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls17
click(Pattern("iste_etabiss.png").similar(0.88))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement17)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement17.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls17.clclick()
click(Pattern("iste_etabiss.png").similar(0.88))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))



#SOC 14018
metadata18 = App.open(metadata18)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata18)
metadata18.close()

xls18 = App.open(xls18)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune18 = App.open(csvCommune18)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls18
click(Pattern("iIidIISJI.png").similar(0.94))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI18 = App.open(csvEPCI18)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls18
click(Pattern("iIidIISJI.png").similar(0.94))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement18)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement18.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls18.clclick()
click(Pattern("iIidIISJI.png").similar(0.94))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14019
metadata19 = App.open(metadata19)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata19)
metadata19.close()

xls19 = App.open(xls19)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune19 = App.open(csvCommune19)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls19
click(Pattern("paces_ass_ma.png").similar(0.82))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI19 = App.open(csvEPCI19)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls19
click(Pattern("paces_ass_ma.png").similar(0.82))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement19)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement19.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls19.clclick()
click(Pattern("paces_ass_ma.png").similar(0.82))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14020
metadata20 = App.open(metadata20)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata20)
metadata20.close()

xls20 = App.open(xls20)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune20 = App.open(csvCommune20)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls20
click(Pattern("IiIidIISJISI.png").similar(0.80))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI20 = App.open(csvEPCI20)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls20
click(Pattern("IiIidIISJISI.png").similar(0.80))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement20)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement20.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls20.clclick()
click(Pattern("IiIidIISJISI.png").similar(0.80))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14021
metadata21 = App.open(metadata21)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata21)
metadata21.close()

xls21 = App.open(xls21)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune21 = App.open(csvCommune21)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls21
click(Pattern("PIaces_ass_f.png").similar(0.87))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI21 = App.open(csvEPCI21)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls21
click(Pattern("PIaces_ass_f.png").similar(0.87))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement21)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement21.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls21.clclick()
click(Pattern("PIaces_ass_f.png").similar(0.87))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14022
metadata22 = App.open(metadata22)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata22)
metadata22.close()

xls22 = App.open(xls22)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune22 = App.open(csvCommune22)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls22
click(Pattern("nbstructJeti.png").similar(0.86))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI22 = App.open(csvEPCI22)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls22
click(Pattern("nbstructJeti.png").similar(0.86))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement22)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement22.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls22.clclick()
click(Pattern("nbstructJeti.png").similar(0.86))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




#SOC 14023
metadata23 = App.open(metadata23)
#Test de verification de presence d'une date de mise a jours; Si faux on passe au Else
if exists(Pattern("0bjtIIaj0bjt.png").similar(0.92)):    
    find(Pattern("0bjtIIaj0bjt.png").similar(0.84))
    hover(Pattern("0bjtIIaj0bjt.png").similar(0.84).targetOffset(-93,-19))

#Fonction maintenir clique gauche
    mouseDown(Button.LEFT)
#Parametre de la fonction    
    mouseMove(Env.getMouseLocation().offset(80, 0))
#Fonction relacher clique gauche    
    mouseUp(Button.LEFT)
#Coller cette date    
    date = input ("Entrer une date de mise a jour (01/01/2013) :")
    print(date)
    paste(date)                 

#Cas second d'execution si aucune date de mise a jour n'est presente
else : 
        find("0bjtIIaj0bjt-1.png")
        hover(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        click(Pattern("0bjtIIaj0bjt-1.png").similar(0.78).targetOffset(-92,-17))
        paste('01/01/2013')

#Test de verification de presence de la balise granularite avec commune; Si faux on passe au Else
if exists("g1anul1i1Cum.png"):      
    find("g1anul1i1Cum.png")
    hover("g1anul1i1Cum.png")  
                
#Cas second d'execution si aucune date de mise a jour n'est presente
else :
        find(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        hover(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        click(Pattern("g1ar1ul1i1.png").similar(0.80).targetOffset(-57,-1))
        #Fonction maintenir clique gauche
        mouseDown(Button.LEFT)
#Parametre de la fonction    
        mouseMove(Env.getMouseLocation().offset(110, 0))
#Fonction relacher clique gauche    
        mouseUp(Button.LEFT)
        paste('<granularite>Commune</granularite>')

type('s', KeyModifier.CTRL)
print(metadata23)
metadata23.close()

xls23 = App.open(xls23)
wait(6)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))

#Valeurs des colonnes/lignes a selectionner et a dupliquer
cellule1 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule1)
paste(cellule1)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))

csvCommune23 = App.open(csvCommune23)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('w', KeyModifier.CTRL)
type(Key.DOWN)
wait(1)
type('v', KeyModifier.CTRL)
wait(1)

type('f', KeyModifier.CTRL)

click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-173,-73))
paste('NC')
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(-91,117))

if exists(Pattern("Rechercherde.png").similar(0.80).targetOffset(-116,-2)):
    click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
    wait(1)
    doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
    doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
    click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
    click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))

elif(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0)):
        click(Pattern("IIIIanIa.png").similar(0.90).targetOffset(-116,0))
        click(Pattern("Rechercherde-1.png").similar(0.90).targetOffset(98,1))
        wait(1)
        doubleClick(Pattern("IEghangerdes.png").similar(0.87).targetOffset(84,2))
        doubleClick(Pattern("giauterdesca.png").similar(0.80).targetOffset(86,4))
        click(Pattern("Cambiner.png").similar(0.80).targetOffset(-24,0))
        click(Pattern("1392212658566.png").similar(0.83).targetOffset(-1,2))


click(Pattern("BRechercherB.png").similar(0.80).targetOffset(136,18))
click(Pattern("BRechercherB.png").similar(0.80).targetOffset(135,117))
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A222')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls23
click(Pattern("nbplacesjnet.png").similar(0.82))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule3 = input ("Entrer une plage de cellule (M2:M22) :")
print(cellule3)
paste(cellule3)
type(Key.ENTER)
type('c', KeyModifier.CTRL)
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(-20,1))


#Ouverture du fichier slice_EPCI 15012
csvEPCI23 = App.open(csvEPCI23)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))

wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
type(Key.LEFT)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
paste('A2:A24')
type(Key.ENTER)
wait(1)
type('c', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type('k', KeyModifier.CTRL)
type('v', KeyModifier.CTRL)
type(Key.RIGHT)
#Dupliquer l'annee
annee = input ("Entrer une annee(2013) :")
print(annee)
paste(annee)
type('h', KeyModifier.CTRL)
find(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
click(Pattern("lA1-1.png").similar(0.92).targetOffset(27,0))
anneeDuplica = input ("Entrer une plage de cellule pour dupliquer l'annee (B2212:B2432) :")
print(anneeDuplica)
paste(anneeDuplica)
type(Key.ENTER)
click(Pattern("Edition.png").similar(0.89).targetOffset(0,1))
hover("Rompiir.png")
click("BasCtrIns.png")
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)

type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0.png").similar(0.80).targetOffset(-96,38))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))


#xls23
click(Pattern("nbplacesjnet.png").similar(0.82))
wait(7)
type('q', KeyModifier.CTRL)
type('h', KeyModifier.CTRL)
find(Pattern("lA1.png").exact().targetOffset(27,1))
hover(Pattern("lA1.png").exact().targetOffset(27,1))
click(Pattern("lA1.png").similar(0.92).targetOffset(27,0))
cellule4 = input ("Entrer une plage de cellules (M2:M222) :")
print(cellule4)
paste(cellule4)
type(Key.ENTER)
type('c', KeyModifier.CTRL)

#Ouverture du fichier slice_departement 15012
App.open(csvDepartement23)
wait(4)
click(Pattern("umpecccident.png").similar(0.80))
click(Pattern("iumpeucciden.png").similar(0.77).targetOffset(133,44))
mouseDown(Button.LEFT)
wait(4)
mouseUp(Button.LEFT)
find(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
click(Pattern("UnicodeUTF8.png").similar(0.86).targetOffset(-29,1))
wait(1)
click(Pattern("1392109154297.png").similar(0.80).targetOffset(1,3))
wait(2)
type('h', KeyModifier.CTRL)
type('b', KeyModifier.CTRL)
type(Key.DOWN)
type(Key.RIGHT)
type(Key.RIGHT)
type('v', KeyModifier.CTRL)
type(Key.LEFT)
annee1 = input ("Entrer une annee (2013) :")
print(annee1)
paste(annee1)
type(Key.LEFT)
paste('LOIRE-ATLANTIQUE')
type('s', KeyModifier.CTRL)
click(Pattern("10pen0ffice0-7.png").similar(0.80).targetOffset(-96,38))
#csvDepartement23.close()
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))
#xls23.clclick()
click(Pattern("nbplacesjnet.png").similar(0.82))
click(Pattern("IClI1llX.png").similar(0.78).targetOffset(26,0))




