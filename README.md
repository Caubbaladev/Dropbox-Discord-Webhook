
# Intégration Dropbox-Discord Webhook

Ce projet utilise l'API de Dropbox pour interagir avec votre compte Dropbox. Suivez les instructions ci-dessous pour configurer correctement l'application et l'environnement.

---

## Prérequis

1. **Créer un compte Developer Dropbox** :

   - Accédez à [Dropbox Developers](https://www.dropbox.com/developers) et connectez-vous avec votre compte Dropbox.
   - Créez une nouvelle application en suivant les étapes du portail.

2. **Récupération des clés d'API** :

   - Une fois l'application créée, notez l'`APP_KEY` et l'`APP_SECRET` dans les informations de l'application.

---

## Obtention du REFRESH\_TOKEN

1. **Générer le code d'autorisation** :

   - Consultez le lien ci-dessous en remplaçant `APPKEYHERE` par votre `APP_KEY` :
     ```
     https://www.dropbox.com/oauth2/authorize?client_id=APPKEYHERE&response_type=code&token_access_type=offline
     ```
   - Cliquez sur **Autoriser** et copiez le code d'autorisation affiché.

2. **Obtenir le REFRESH\_TOKEN** :

   - Exécutez la commande suivante dans un terminal en remplaçant les valeurs :
     ```bash
     curl https://api.dropbox.com/oauth2/token \
         -d code=AUTHORIZATIONCODEHERE \
         -d grant_type=authorization_code \
         -u APPKEYHERE:APPSECRETHERE
     ```
   - Le retour de cette commande contiendra votre `REFRESH_TOKEN`. Notez-le pour une utilisation future.
   - Le format général de réponse pour le `REFRESH_TOKEN` est le suivant :
     ```json
     {
       "access_token": "ACCESS_TOKEN_VALUE",
       "token_type": "bearer",
       "expires_in": 14400,
       "refresh_token": "REFRESH_TOKEN_VALUE",
       "scope": "account_info.read file_requests.read files.content.read files.content.write files.metadata.read files.metadata.write",
       "uid": "USER_ID",
       "account_id": "ACCOUNT_ID"
     }
     ```

---

## Configuration du code (TRES IMPORTANT !)

1. **Ajout des clés au code** :
   - Assurez-vous que votre `APP_KEY` et `APP_SECRET` sont correctement référencés dans le code source.
   - Ces informations se trouvent dans le portail Developer Dropbox.

2. **Champs à personnaliser dans le code (fichier main.py)** :
   - Remplacez APP_KEY, APP_SECRET, REFRESH_TOKEN, et DISCORD_WEBHOOK_URL par vos propres valeurs.
   - Définissez également MAIN_FILE = 'MPSI --ANNEE--', par exemple, MAIN_FILE = 'MPSI 23-24' en utilisant le nom du fichier spécifique à votre année dans le dossier Dropbox de physique.
  
3. **Initialisation des fichiers docx** :
   - Les fichiers SAMPLE_FILE.docx et UPDATED_FILE.docx doivent être initialisés en tant que copies exactes du fichier Quoi de neuf.docx de la Dropbox de votre classe de MPSI. Ces fichiers doivent contenir exactement le même contenu que l'original.

---

## Consignes spéciales pour la compression

1. **Format correct de compression** :
   - Lorsque vous compressez les fichiers pour une utilisation ou un déploiement, incluez uniquement les fichiers suivants dans l'archive `.zip` :
     - `main.py`
     - `requirements.txt`
     - Fichiers `.docx`
   - **Ne pas inclure de dossiers** directement dans l'archive.

2. **Création d'une fonction dans Google Cloud Functions** :
   - Avant d'importer l'archive `.zip`, créez une nouvelle fonction dans Google Cloud Functions.
   - Choisissez l'environnement d'exécution approprié (par exemple, Python 3.10).

3. **Téléchargement et déploiement** :
   - Accédez à l'interface Google Cloud Functions.
   - Dans la section "Code", sélectionnez "Importer un fichier ZIP".
   - Téléchargez l'archive `.zip` contenant `main.py`, `requirements.txt`, et vos fichiers `.docx`.


---

## Notes supplémentaires

- Veillez à garder vos clés d'application et tokens confidentiels.
- Consultez la documentation officielle de l'API Dropbox pour plus de détails : [API Dropbox](https://www.dropbox.com/developers/documentation)

---

Merci d'utiliser ce guide pour une configuration réussie de votre application Dropbox.

