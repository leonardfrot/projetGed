using System;
using Newtonsoft.Json.Linq;
using RestSharp;
using System.Collections.Generic;
using Aspose.Cells;
using UnityEngine;
using Aspose.Cells.Utility;
using System.IO;

class User
{
    public User(string firstname, string lastname)
    {
        this.firstname = firstname;
        this.lastname = lastname;
    }
    public string firstname;
    public string lastname;
}




namespace GED_projet
{
    class Program
    {
        static void Main(string[] args)
        {
            // connection
            String url = "http://157.26.82.44:2240/token";
            String username = "Leonard";
            String password = "Leonard";
            String token = GetToken(url, username, password);
            Console.WriteLine("vous êtes connecté");

            //get userinformation and welcome
            User user = getUserInformation(token);

            Console.WriteLine("bienvenue " + user.firstname + " " + user.lastname);

            

            while(true)
            {


                //sendFileToDigitalCorner(token, "exemple_facture.pdf", "Gestion des factures 2");


                exportMetaDataInCSV(token, "Gestion des factures 2");

                Console.WriteLine("q for quit");

                String res = Console.ReadLine();

                if(res == "q")
                {
                    break;
                }


            }
        }

        // cette méthode retour le json avec le tocken.
        static string GetToken(String url, String username, String password)
        {
            var client = new RestClient(url);
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("Content-Type", "application/x-www-form-urlencoded");
            request.AddParameter("grant_type", "password");
            request.AddParameter("username", username);
            request.AddParameter("password", password);
            IRestResponse response = client.Execute(request);
            String jsonResponse = response.Content.ToString();
            JObject jObject = Newtonsoft.Json.Linq.JObject.Parse(jsonResponse);
            var token = jObject.SelectToken("access_token");
            return token.ToString();
        }

        // cette méthode est appelé pour trouver les informations pour l'utilisateur
        static User getUserInformation(String token)
        {
            var client = new RestClient("http://157.26.82.44:2240/API/account/current");
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Bearer " + token);
            IRestResponse response = client.Execute(request);
            String jsonResponse = response.Content.ToString();
            JObject jObject = Newtonsoft.Json.Linq.JObject.Parse(jsonResponse);
            string firstname = (string)jObject["Name"];
            string lastname = (string)jObject["Surname"];
            User user = new User(firstname, lastname);
            return user;
        }



        // Cette méthode est appelé pour récupérer la structure de la facture
        static string getInvoiceStructure(String token, String contentTypeName)
        {
            var client = new RestClient("http://157.26.82.44:2240/api/content-type/list");
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Bearer " + token);
            IRestResponse response = client.Execute(request);
            String jsonResponse = response.Content.ToString();
            JArray jArray = Newtonsoft.Json.Linq.JArray.Parse(jsonResponse);

            String contentTypeId = null;

            for (int i = 0; i < jArray.Count; i++)
            {
                if (jArray[i]["text"].ToString() == contentTypeName)
                {
                    contentTypeId = jArray[i]["id"].ToString();
                    break;
                }
            }

            var client2 = new RestClient("http://157.26.82.44:2240/api/document/structure/" + contentTypeId);
            client2.Timeout = -1;
            var request2 = new RestRequest(Method.GET);
            request2.AddHeader("Authorization", "Bearer " + token);
            IRestResponse response2 = client2.Execute(request);

            

            return response2.Content.ToString();
        }


        // la méthode pour transofrmer en base 64
        static String TransformPDFToBase64(String filepath)
        {
            byte[] bytes = System.IO.File.ReadAllBytes(filepath);

            System.IO.File.WriteAllBytes(filepath, bytes);
            String fileBase64 = Convert.ToBase64String(bytes);
            return fileBase64;

        }

        // la méthode pour envoyer la facture
        static void sendFileToDigitalCorner(String token, String filepath, String contentTypeName)
        {
            // on récupère le fichier en base 64
            String fileBase64 = TransformPDFToBase64(filepath);
            // on récupère la structure json
            String jsonStructure = getInvoiceStructure(token, contentTypeName);
            // on modifie la jsonStructure pour envoyer la nouvelle facture
            JObject rss = JObject.Parse(jsonStructure);

            //créer un nouvel Jobjet à envoyer d'abord pour les attachement
            var attachement = new JObject();
            attachement.Add("id", 0);
            attachement.Add("fileName", filepath);
            attachement.Add("base64File", fileBase64);

            // construction des fe
            var fields = new JArray();
            JArray fielsFromStructure = (JArray)rss["Fields"];
            


            List<String> values = new List<String>();

            //ff date
            values.Add("20.12.2021");
            // fournisseur
            values.Add("Tornos");
            // num facutre
            values.Add("123444");
            // description
            values.Add("pas de description");
            //etat
            values.Add("A editer");
            //"responsable"
            values.Add("Admin Groupe 2");
            // "sous-validateur"
            values.Add("Admin Groupe 2");
            // validateur1
            values.Add("Autre 1");
            // validateur2
            values.Add("compta 2");
            //direction générale
            values.Add("Vincent Chassot");
            // comptabilité
            values.Add("Mounir Marzouk");
            // date de l'échéance
            values.Add("19.01.2022");
            //swift
            values.Add(null);
            // numero de compte
            values.Add(null);
            //devise facture
            values.Add("CHF");
            // rabais
            values.Add(null);
            // escompte
            values.Add(null);
            // TVA
            values.Add("Normal");
            // Montant HT
            values.Add(null);
            // Montant TTC
            values.Add(null);




            for (int i = 0; i < fielsFromStructure.Count; i++)
            {
                var oneFields = new JObject();
                oneFields.Add("code", fielsFromStructure[i]["Code"]);
                oneFields.Add("value", values[i]);
                fields.Add(oneFields);

            }

            

            var mainJson = new JObject();
            mainJson.Add("objectID", 0);
            mainJson.Add("ContentTypeID", rss["ContentTypeID"]);
            mainJson.Add("fields", fields);
            mainJson.Add("attachment", attachement);

            var client = new RestClient("http://157.26.82.44:2240/api/document/save");
            client.Timeout = -1;
            var request = new RestRequest(Method.POST);
            request.AddHeader("Authorization", "Bearer " + token);
            request.AddHeader("Content-Type", "application/json");

            request.AddParameter("application/json", mainJson, RestSharp.ParameterType.RequestBody);
            IRestResponse response = client.Execute(request);
            Console.WriteLine("file sent");
        }




        static List<string> getAllDocumentId(String token, String contentTypeName)
        {


            List<string> idList = new List<string>();

            // récupération de l'id du content type.

            var client = new RestClient("http://157.26.82.44:2240/api/content-type/list");
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Bearer " + token);
            IRestResponse response = client.Execute(request);

            String jsonResponse = response.Content.ToString();
            JArray jArray = Newtonsoft.Json.Linq.JArray.Parse(jsonResponse);



            String contentTypeId = null;

            for (int i = 0; i < jArray.Count; i++)
            {
                if (jArray[i]["text"].ToString() == contentTypeName)
                {
                    contentTypeId = jArray[i]["id"].ToString();
                    break;
                }
            };
            

             // construction du search pattern à envoyer

            String searchPattern = "FF_2_ETAT|l01|en attente d'export|list";

            var mainJson = new JObject();
            mainJson.Add("searchPattern", searchPattern);
            mainJson.Add("ContentTypeID", contentTypeId);

            // recherche avancé pour récupérer la liste des objectId

            var client2 = new RestClient("http://157.26.82.44:2240/api/search/advanced");
            client2.Timeout = -1;
            var request2 = new RestRequest(Method.POST);
            request2.AddHeader("Authorization", "Bearer " + token);
            request2.AddHeader("Content-Type", "application/json");
            var body = mainJson;
            request2.AddParameter("application/json", body, RestSharp.ParameterType.RequestBody);
            IRestResponse response2 = client2.Execute(request2);

            String jsonResponse2 = response2.Content.ToString();
            JArray jArray2 = Newtonsoft.Json.Linq.JArray.Parse(jsonResponse2);

            for (int i = 0; i < jArray2.Count; i++)
            {
                idList.Add(jArray2[i]["ObjectID"].ToString());
            };

            
            

            return idList;
        }

        static void exportPDF(String fileToTransform, String fileName)
        {
            byte[] sPDFDecoded = Convert.FromBase64String(fileToTransform);

            string[] arr = fileName.Split('.');

            string newfilename = arr[0] + "_exported.pdf";


            File.WriteAllBytes(newfilename, sPDFDecoded);
            Console.WriteLine("pdf exported");

        }


        static void exportMetaDataInCSV(String token, String contentTypeName)
        {
            List<string> idList = getAllDocumentId(token, contentTypeName);

            var mainJson = new JArray();

            

            for (int i = 0; i < idList.Count; i++)
            {
                
                // récupération des métadonnées
                var client = new RestClient("http://157.26.82.44:2240/api/document/" + idList[i] + "/metadata");
                client.Timeout = -1;
                var request = new RestRequest(Method.GET);
                request.AddHeader("Authorization", "Bearer " + token);
                IRestResponse response = client.Execute(request);

                String jsonResponse = response.Content.ToString();


                // construire un jArrayAvec tous les documents
                JObject jObject = Newtonsoft.Json.Linq.JObject.Parse(jsonResponse);

                mainJson.Add(jObject);

                JArray jArray = (JArray)jObject["Fields"];

                int size = jArray.Count;
                
                
                for (int j = 0; j < size; j++)
                {
                    mainJson.Add(null);
                };

                // récupération des attachments


                var client2 = new RestClient("http://157.26.82.44:2240/api/document/" + idList[i]+ "/attachment");
                client2.Timeout = -1;
                var request2 = new RestRequest(Method.GET);
                request2.AddHeader("Authorization", "Bearer " + token);
                IRestResponse response2 = client2.Execute(request2);
               
                String jsonResponse2 = response2.Content.ToString();
                JObject jObject2 = Newtonsoft.Json.Linq.JObject.Parse(jsonResponse2);
                String filename = jObject2["FileName"].ToString();
                
                String fileToTransform = jObject2["File"].ToString();

                exportPDF(fileToTransform, filename);





            }

            string jsonInput = mainJson.ToString();
            

            var workbook = new Workbook();

            var worksheet = workbook.Worksheets[0];

            var layoutOptions = new JsonLayoutOptions();
            layoutOptions.ArrayAsTable = true;

            JsonUtility.ImportData(jsonInput, worksheet.Cells, 0, 0, layoutOptions);

            workbook.Save("output.csv", SaveFormat.CSV);
            Console.WriteLine("csv exported");


        }

       

    }
}







