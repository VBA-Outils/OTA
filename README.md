# OTA
<h1>Licence</h1>
<p>Ce projet est distribué sous licence MIT. Consultez le fichier LICENSE pour plus de détails.</p>
<h1>Prérequis</h1>
<p>Environnement de développement : Microsoft Visual Basic for Applications (VBA)</p>
<h1>Reformater le résultat d'une requête pour les OTA</h1>
<h2>Installation</h2>
<p>Importer le module "PressePapiers.bas" dans le projet VBA.<p>
<p>Cliquer sur le menu "Développeur", puis "Insérer", "Contrôle de formulaire", "Bouton". Placer le bouton sur la feuille souhaitée et lui associer la macro adéquate : OTASquirrel par exemple.<p>
</p>Copier le résultat de la requête dans le presse-papiers, puis cliquer sur le bouton créé précédement, récupérer les informations du presse-papiers Windows en les collant directement dans l'OTA.</p>
<h2>Fonctionnalité</h2>
<p>Lors de la création d'OTA, les résultats de requêtes peuvent être comparés à des valeurs de référence afin de vérifier la non régression de l'applicatif.</p>
<p>OTA exécute la requête du script, puis compare le résultat obtenu avec les valeurs attendues. Cette liste de valeur doit avoir le formattage suivant : |Donnée 1|Donnée 2| où chaque résultat attendu est séparé par la barre verticale.</p>
<p>Les macros OTApgAdmin et OTAsquirrel permettent de reformatter le résultat d'une requête exécutée dans respectivement PgAdmin et Squirrel afin de pouvoir coller directement dans l'OTA le résultat attendu.</p>
<h2>Méthodologie</h2>
<p>Exécuter la requête de l'OTA dans l'outil de requêtage (Squirrel, Pgadmin, etc), puis sélectionner toutes les valeurs du résultat obtenu. Copier ce résultat dans le presse-papiers, exécuter la macro adéquate, coller dans OTA le nouveau résultat reformatté.</p>
<h3>Exemple</h3>
<p>Résultat de la requête</p>
<table><tr><th>t_ref_cd_departement</th><th>t_ref_lib_departement</th><th>t_ref_nb_hab</th></tr><tr><td>01</td><td>Ain</td><td>671&nbsp;289</td></tr><tr><td>02</td><td>Aisne</td><td>525&nbsp;558</td></tr></table>
<p>Résultat après exécution de la macro</p>
<p>
|t_ref_cd_departement|t_ref_lib_departement|t_ref_nb_hab|<br>
|01|Ain|352000|<br>
|02|Aisne|452365|
</p>
