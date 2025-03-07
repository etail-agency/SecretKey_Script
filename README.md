# Script de Gestion des Clés Secrètes Amazon

Ce script permet d'automatiser la gestion des clés secrètes pour les comptes Amazon Vendor Central. Il vérifie les dates d'expiration des clés et les renouvelle automatiquement si nécessaire.

## Prérequis

- Python 3.8 ou supérieur
- Chrome WebDriver
- Compte Amazon Vendor Central avec les permissions nécessaires

## Installation

1. Clonez ce dépôt :
```bash
git clone [URL_DU_REPO]
cd Script_client_SectretKey
```

2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

3. Configurez les variables d'environnement :
   - Copiez le fichier `.env.example` vers `.env`
   - Remplissez les informations dans le fichier `.env`

## Configuration

1. Assurez-vous que le fichier Excel `listAccount_sorted.xlsx` est présent dans le répertoire du projet
2. Le fichier Excel doit contenir deux colonnes :
   - AccountName : Nom du compte
   - ClientID : ID du client

## Utilisation

Pour exécuter le script :

```bash
python clientSecret.py
```

Le script va :
1. Se connecter à Amazon Vendor Central
2. Parcourir la liste des comptes
3. Vérifier les dates d'expiration des clés
4. Renouveler les clés si nécessaire
5. Générer un rapport Excel avec les résultats

## Structure des Fichiers

- `clientSecret.py` : Script principal
- `listAccount_sorted.xlsx` : Liste des comptes à traiter
- `.env` : Fichier de configuration (à créer)
- `requirements.txt` : Liste des dépendances

## Sécurité

- Ne partagez jamais vos clés d'API ou vos identifiants
- Gardez le fichier `.env` sécurisé et ne le partagez pas
- Ne committez pas le fichier `.env` dans le dépôt Git

## Support

Pour toute question ou problème, veuillez créer une issue dans le dépôt GitHub. 