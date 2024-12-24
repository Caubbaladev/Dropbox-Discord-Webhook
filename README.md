
# Dropbox Integration Guide

Ce projet utilise l'API de Dropbox pour interagir avec votre compte Dropbox. Suivez les instructions ci-dessous pour configurer correctement l'application et l'environnement.

---

## Prérequis

1. **Créer un compte Developer Dropbox** :

   - Accédez à [Dropbox Developers](https://www.dropbox.com/developers) et connectez-vous avec votre compte Dropbox.
   - Créez une nouvelle application en suivant les étapes du portail.

2. **Récupération des clés d'API** :

   - Une fois l'application créée, notez l'`APP_KEY` et l'`APP_SECRET` dans les informations de l'application.

---

## Obtention du REFRESH_TOKEN

1. **Générer le code d'autorisation** :

   - Consultez le lien ci-dessous en remplaçant `APPKEYHERE` par votre `APP_KEY` :
     ```
     https://www.dropbox.com/oauth2/authorize?client_id=APPKEYHERE&response_type=code&token_access_type=offline
     ```
   - Cliquez sur **Autoriser** et copiez le code d'autorisation affiché.

2. **Obtenir le REFRESH_TOKEN** :

   - Exécutez la commande suivante dans un terminal en remplaçant les valeurs :
     ```bash
     curl https://api.dropbox.com/oauth2/token          -d code=AUTHORIZATIONCODEHERE          -d grant_type=authorization_code          -u APPKEYHERE:APPSECRETHERE
     ```
   - Le retour de cette commande contiendra votre `REFRESH_TOKEN`. Notez-le pour une utilisation future.

---

## Configuration du code

1. **Ajout des clés au code** :
   - Assurez-vous que votre `APP_KEY` et `APP_SECRET` sont correctement référencés dans le code source.
   - Ces informations se trouvent dans le portail Developer Dropbox.

---

## Consignes spéciales pour la compression

1. **Format correct de compression** :
   - Lorsque vous compressez les fichiers pour une utilisation ou un déploiement, incluez uniquement les fichiers suivants dans l'archive `.zip` :
     - `main.py`
     - `requirements.txt`
     - Fichiers `.docx`
   - **Ne pas inclure de dossiers** directement dans l'archive.

---

## Notes supplémentaires

- Veillez à garder vos clés d'application et tokens confidentiels.
- Consultez la documentation officielle de l'API Dropbox pour plus de détails : [API Dropbox](https://www.dropbox.com/developers/documentation)

---

Merci d'utiliser ce guide pour une configuration réussie de votre application Dropbox.
