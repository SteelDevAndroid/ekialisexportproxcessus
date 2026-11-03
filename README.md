# Export des processus

Application en ligne de commande **Java 25 / Spring Boot 4.2** qui complète la
première feuille d'un classeur `.xlsx` avec les colonnes `pr`, `sca`, `pr1`, `pr2`
et `pr3`, via Apache POI.

Le projet utilise **Spring Boot 4.2.0-RC1**, version 4.2 disponible sur Maven Central
au moment de sa création (la version finale 4.2.0 n'est pas encore publiée).

## Compiler et exécuter

Prérequis : JDK 25 et Maven 3.9 ou supérieur.

```sh
mvn clean verify
java -jar target/export-processus-1.0.0.jar
```

Par défaut, le programme lit `process.xlsx` dans le répertoire courant et crée
`process-pr.xlsx` dans le même répertoire. Pour choisir les chemins :

```sh
java -jar target/export-processus-1.0.0.jar --input=/chemin/process.xlsx --output=/chemin/resultat.xlsx
```

Les chemins contenant des espaces doivent être placés entre guillemets.
Le répertoire de sortie doit exister. Le programme s'arrête après le traitement.

## Intégration continue

Le workflow GitHub Actions **Build JAR** lance les tests et génère le JAR avec
Java 25 à chaque pull request et à chaque push sur `main`. Il peut aussi être
lancé manuellement depuis l'onglet **Actions** une fois présent sur `main`.

Après une exécution réussie, ouvrir **Actions → Build JAR → l'exécution → Artifacts**
et télécharger **export-processus-jar**. L'archive contient le JAR exécutable
`export-processus-1.0.0.jar` ; elle est conservée pendant 14 jours. Extraire le JAR,
puis l'exécuter avec Java 25 selon la commande ci-dessus.

## Règles de traitement

La première ligne doit contenir les en-têtes `Niveau` et `Processus` ; leur casse,
leur position et les espaces entourant les en-têtes n'ont pas d'importance.
Les niveaux peuvent être des nombres Excel ou du texte (`1`, `2`, `3`).
L'ordre des lignes et les libellés sont conservés.

| Niveau | Processus | pr |
| --- | --- | --- |
| 1 | A | A |
| 2 | B | A#B |
| 3 | C | A#B#C |
| 3 | D | A#B#D |
| 2 | E | A#E |
| 3 | F | A#E#F |
| 1 | G | G |

Chaque niveau 1 réinitialise le parent de niveau 2. Les lignes entièrement vides
sont ignorées sans réinitialiser les parents. Une hiérarchie incomplète (par
exemple, un niveau 3 directement après un nouveau niveau 1), un niveau invalide ou
un libellé vide provoque une erreur indiquant le numéro de ligne Excel.

Les colonnes `pr`, `sca`, `pr1`, `pr2` et `pr3` sont ajoutées après les colonnes
existantes. Si elles existent déjà, elles sont recalculées. Les autres cellules,
leurs styles et les autres feuilles sont
conservés. Le fichier source reste inchangé ; un fichier de sortie existant n'est
jamais écrasé. Le résultat n'est écrit qu'après validation de toutes les lignes.

### Société et décomposition de `pr`

`sca` est calculée à partir du libellé complet `pr`, parents compris. La recherche
est sensible à la casse et vérifie la présence des chaînes suivantes dans cet
ordre : `ASF` → `ASF`, `ESC` ou `ESCOTA` → `ESC`, `COF` ou `COFIROUTE` → `COF`,
`VA` → `VA`. Si plusieurs codes sont présents, le premier de cet ordre est retenu.
Sans correspondance, `sca` est vide.

Pour la décomposition, le préfixe initial exact `VA - `, `ESC - `, `COF - ` ou
`ASF - ` est retiré. Aucun autre préfixe n'est retiré, et `pr` reste inchangé.
Le texte restant est découpé aux deux premiers séparateurs ` - ` ou `#`, dans
l'ordre de leur apparition, sans inclure ces séparateurs :

- `pr1` contient la partie avant le premier séparateur, ou tout le texte sans séparateur.
- `pr2` contient la partie entre les deux premiers séparateurs, ou le reste après un seul séparateur.
- `pr3` contient tout le reste après le deuxième séparateur, y compris les séparateurs suivants.

Les parties absentes ou vides sont écrites comme des chaînes vides. Les autres
espaces et les préfixes présents au milieu du libellé sont conservés.

| pr | sca | pr1 | pr2 | pr3 |
| --- | --- | --- | --- | --- |
| ASF - Assister - Appeler - Informer | ASF | Assister | Appeler | Informer |
| ESC - Assister#Appeler#Informer#Archiver | ESC | Assister | Appeler | Informer#Archiver |
| COFIROUTE - Assister | COF | COFIROUTE | Assister | |
| VA - Assister | VA | Assister | | |
| Assister | | Assister | | |

## Tests

```sh
mvn test
```

Les tests utilisent de vrais classeurs Excel temporaires et vérifient les trois
niveaux, les changements de parent, les lignes vides, les entrées invalides, le
recalcul des colonnes générées, les codes de société, les préfixes, les séparateurs
mixtes et les parties vides, la conservation des formules/styles/feuilles et la
protection contre l'écrasement des fichiers.
