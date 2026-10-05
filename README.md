# Smart Mail Assistant

J'ai demandé a claude de reflcechir sur le developpement d'une app de gestion de mail assistée par IA. Peux tu me dire ce que tu en penses dans un premier temps sans rien développer et me dire s'il est plus intéressant que cette app tourne en local sur un pc ou s'il serait intéressant que l'app soit une sorte de site web qui permet de gérer a distance sa boite mail  ?

Le fichier ci-dessus contient maintenant un assistant de configuration manuelle, sans plus aucune variable d'environnement à régler.

Au premier lancement, une page demande :

l'adresse e-mail, l'identifiant de connexion (souvent identique, mais pas toujours) et le mot de passe ;

le serveur IMAP, son port et sa sécurité (SSL/TLS, STARTTLS ou aucune) ;

le serveur SMTP, son port et sa sécurité ;

la clé API et la signature des réponses.

Dès que l'adresse est saisie, l'assistant préremplit des valeurs probables (imap.domaine, smtp.domaine, ports 993 et 587), à corriger selon ce que donne le service informatique ou le PO. Un bouton Tester la connexion vérifie séparément la réception et l'envoi, et indique laquelle des deux échoue. Le bouton ⚙ permet de rouvrir ces réglages plus tard, et un mot de passe laissé vide conserve l'ancien.

Limites actuelles

La configuration est stockée dans ~/.mailpilot/config.json, en clair, avec des droits restreints au seul utilisateur. Pour une diffusion en école, il faudra passer par le coffre du système (Windows Credential Manager, Trousseau Mac).

Certains serveurs de PO ont des certificats internes que Python refuse. Il faudra alors prévoir une option « faire confiance à ce certificat ».

Un domaine comme enseignementbw.be peut très bien être hébergé derrière Microsoft 365 ou Google, même s'il a son nom propre. Dans ce cas l'accès IMAP par mot de passe peut être bloqué, et il faudra l'OAuth que j'évoquais. Le bouton de test le révélera tout de suite.

Suite logique : une détection automatique des paramètres (les domaines publient souvent leur configuration mail), puis les fiches par réseau (PO du BW, WBE, SEGEC) avec des valeurs types prêtes à l'emploi, et enfin l'emballage en .exe. Si tu me donnes les paramètres IMAP/SMTP de deux ou trois de ces réseaux, je peux déjà préparer ces fiches.

This project was built with [Lovable](https://lovable.dev).

## Build with Lovable

Continue developing this project in the [Lovable editor](https://lovable.dev/projects/be836fce-e973-4de1-9fa0-d25c0f8b0dca).

- **Ship faster**: describe what you want to build and Lovable handles the code.
- **Stay in sync**: every change made in Lovable is committed straight to this repository.
- **Full ownership**: this code is yours. Push to `main` on GitHub and your changes sync back into Lovable, ready for your next prompt.

## Development

Prefer working locally? You need Node.js and npm — [install with nvm](https://github.com/nvm-sh/nvm#installing-and-updating).

```sh
git clone <this-repository-url>
cd <repository-name>
npm i
npm run dev
```
