# Heyxcel - Librairie Excel

## Introduction
    Librairie pour lire les données d'un fichier excel.

## Lecture

```cs
Heyxcel clientExcel = new Heyxcel(@"C:\Users\ldamiens\Desktop\DataHappyAuto\CLIENT.xls");
ClientBL clientBL = new ClientBL();
int start = 2;
int end = 9999;
try{
    clientExcel.Open();
    clientExcel.Read("A", start, end);
    clientExcel.Read("C", start, end);
    clientExcel.Read("H", start, end);
    clientExcel.Close();
}
catch(Exception exception){
    Console.WriteLine($"Exception : {exception}");
}

// Exemple pour récupérer les données stockées.

Dictionary<string, string> ligne22 = clientExcel.ReadRowFromStoredValues(22);
Dictionary<string, string> ligne33 = clientExcel.ReadRowFromStoredValues(33);
Dictionary<string, string> ligne44 = clientExcel.ReadRowFromStoredValues(44);

// Exemple pour récupérer toutes les données.

int echec = 0;

for(int i = start; i <= end; i++){
    Client client = new Client();
    Dictionnary<string, string> ligne = clientExcel.ReadRowFromStoredValues(i);
    client.Map(ligne);
    
    //On passe par le business logic et on créer le nouveau client.
    if(!ClientBL.Create(client)){
        Console.WriteLine($"Erreur lors de la création du client n°{client.idClient}");
        echec++;
    }
    Console.Title = $"Insertion client - {i}/{end} || Echoué : {echec}";
}
```
- On créer une nouvelle instance du lecteur de fichier excel puis on lui passe en paramètre le chemin absolue du fichier.
- On déclare un nouveau BusinessLogic de l'objet que l'on souhaite manipuler.
- On définit les deux bornes de début et fin de ligne du fichier excel.
- On ouvre le fichier et on lit les colonnes que l'on souhaite, elles seront stockées en mémoire.
- On récupère les lignes que l'on souhaite OU on les récupères toutes via une boucle for
- On créer un nouveau modèle de l'objet que l'on souhaite manipuler
- On lit la ligne que l'on veut
- On map les données stockées dans notre nouveau modèle
- On passe par le Business Logic pour créer une nouvelle ligne en base
- On actualise le status du programme via le titre de la console

## Ecriture

    Pas encore documenté.