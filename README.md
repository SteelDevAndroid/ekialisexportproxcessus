# Export des processus

Application en ligne de commande **Java 25 / Spring Boot 4.2** qui complète la
première feuille d'un classeur `.xlsx` avec la colonne `pr`, via Apache POI.

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

La colonne `pr` est ajoutée après les colonnes existantes. Si elle existe déjà,
elle est recalculée. Les autres cellules, leurs styles et les autres feuilles sont
conservés. Le fichier source reste inchangé ; un fichier de sortie existant n'est
jamais écrasé. Le résultat n'est écrit qu'après validation de toutes les lignes.

## Tests

```sh
mvn test
```

Les tests utilisent de vrais classeurs Excel temporaires et vérifient les trois
niveaux, les changements de parent, les lignes vides, les entrées invalides, le
recalcul de `pr`, la conservation des formules/styles/feuilles et la protection
contre l'écrasement des fichiers.
