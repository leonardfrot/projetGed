using System;
using Newtonsoft.Json.Linq;
using RestSharp;
using System.Collections.Generic;
using Aspose.Cells;
using UnityEngine;
using Aspose.Cells.Utility;
using System.IO;
using System.Xml;


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


public class Facture
{
    public string Date { get; set; }
    public string Fournisseur { get; set; }
    public string NumFacture { get; set; }
    public string Description { get; set; }
    public string Etat { get; set; }
    public string Responsable { get; set; }
    public string SousValidateur { get; set; }
    public string Validateur1 { get; set; }
    public string Validateur2 { get; set; }
    public string DirectionGenerale { get; set; }
    
    public string DateEcheance { get; set; }
    
    public string NumCompte { get; set; }
    public string Devise { get; set; }
    
    public string Tva { get; set; }
    
    public string MontantTTC { get; set; }

    public string Comptabilite { get; set; }

    public string Rabais { get; set; }
    public string Escompte { get; set; }

    public string MontantHT { get; set; }

    public string Swift { get; set; }

    

    


    public Facture(string date, string fournisseur, string numFacture, string description,
    string etat, string responsable, string sousValidateur, string validateur1, string validateur2,
    string directionGenerale, string comptabilite, string dateEcheance, string swift,
    string numCompte, string devise, string rabais, string escompte, string tva,
    string montantHT, string montantTTC)
    {
        Date = date;
        Fournisseur = fournisseur;
        NumFacture = numFacture;
        Description = description;
        Etat = etat;
        Responsable = responsable;
        SousValidateur = sousValidateur;
        Validateur1 = validateur1;
        Validateur2 = validateur2;
        DirectionGenerale = directionGenerale;
        Comptabilite = comptabilite;
        DateEcheance = dateEcheance;
        Swift = swift;
        NumCompte = numCompte;
        Devise = devise;
        Rabais = rabais;
        Escompte = escompte;
        Tva = tva;
        MontantHT = montantHT;
        MontantTTC = montantTTC;
    }
    public string getValue(int code)
    {
        switch (code)
        {
            case 0:
                return Date;
            case 1:
                return Fournisseur;
            case 2:
                return NumFacture;
            case 3:
                return Description;
            case 4:
                return Etat;
            case 5:
                return Responsable;
            case 6:
                return SousValidateur;
            case 7:
                return Validateur1;
            case 8:
                return Validateur2;
            case 9:
                return DirectionGenerale;
            case 10:
                return DateEcheance;
            case 11:
                return NumCompte;
            case 12:
                return Devise;
            case 13:
                return Tva;
            case 14:
                return MontantTTC;
            case 15:
                return Comptabilite;
            case 16:
                return Rabais;
            case 17:
                return Escompte;
            case 18:
                return MontantHT;
            case 19:
                return Swift;
            default:
                return null;
        }
    }
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

            //getuserinformation and welcome
            User user = getUserInformation(token);

            Console.WriteLine("bienvenue " + user.firstname + " " + user.lastname);

            

            while(true)
            {

                String filepath = "C:\\digital_corner\\Import\\exemple_facture_2.pdf";
                String metaDataXML = "C:\\digital_corner\\Import\\metadata_facture1.xml";
                String type = "Facture_fourn";

                //sendFileToDigitalCorner(token, filepath, type, metaDataXML);


                exportMetaDataInCSV(token, type);

                Console.WriteLine("q for quit");

                String res = Console.ReadLine();

                if(res == "q")
                {
                    break;
                }


            }
        }

        //cette méthode est pour lire un xml de metadonnées
        static List<Facture> ReadXML(String metadata)
        {
            XmlDocument xmlDocument = new XmlDocument();

            List<Facture> factures = new List<Facture>();

            xmlDocument.Load(metadata);

            XmlNodeList xmlNodeList = xmlDocument.SelectNodes("/Factures/Facture");

            Facture facturee;
            

            foreach (XmlNode xmlNode in xmlNodeList)
            {

                
                facturee = new Facture(xmlNode["Date"].InnerText, xmlNode["Fournisseur"].InnerText, xmlNode["NumFacture"].InnerText,
                    xmlNode["Description"].InnerText, xmlNode["Etat"].InnerText, xmlNode["Responsable"].InnerText,
                    xmlNode["SousValidateur"].InnerText, xmlNode["Validateur1"].InnerText, xmlNode["Validateur2"].InnerText,
                    xmlNode["DirectionGenerale"].InnerText, xmlNode["Comptabilite"].InnerText, xmlNode["DateEcheance"].InnerText,
                    xmlNode["Swift"].InnerText, xmlNode["NumCompte"].InnerText, xmlNode["Devise"].InnerText, xmlNode["Rabais"].InnerText,
                    xmlNode["Escompte"].InnerText, xmlNode["Tva"].InnerText, xmlNode["MontantHT"].InnerText, xmlNode["MontantTTC"].InnerText);
                    factures.Add(facturee);
            }

            
            return factures;
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
        static void sendFileToDigitalCorner(String token, String filepath, String contentTypeName, String metadata)
        {
            // on récupère le fichier en base 64
            String fileBase64 = TransformPDFToBase64(filepath);
            // on récupère la structure json
            String jsonStructure = getInvoiceStructure(token, contentTypeName);
            // on modifie la jsonStructure pour envoyer la nouvelle facture
            JObject rss = JObject.Parse(jsonStructure);
            

            //créer un nouvel Jobjet à envoyer d'abord pour les attachement

            String[] filenameArr = filepath.Split('\\');
            String filename = filenameArr[3];
            
            var attachement = new JObject();
            attachement.Add("id", 0);
            attachement.Add("fileName", filename);
            attachement.Add("base64File", fileBase64);

            // construction des fe
            var fields = new JArray();
            JArray fielsFromStructure = (JArray)rss["Fields"];


            Facture facture = ReadXML(metadata)[0];

            for (int i = 0; i < fielsFromStructure.Count; i++)
            {
                var oneFields = new JObject();
                oneFields.Add("code", fielsFromStructure[i]["Code"]);
                oneFields.Add("value", facture.getValue(i));
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
            
            Console.WriteLine(filepath + " sent");
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


        static void changeState(String token,  String id)
        {
            var client = new RestClient("http://157.26.82.44:2240/api/document/"+id+"/metadata");
            client.Timeout = -1;
            var request = new RestRequest(Method.GET);
            request.AddHeader("Authorization", "Bearer " + token);
            IRestResponse response = client.Execute(request);
            
            String jsonResponse = response.Content.ToString();
            JObject jobject = Newtonsoft.Json.Linq.JObject.Parse(jsonResponse);
            jobject["Fields"][4]["Value"] = "Archivé";


            
            

            var client2 = new RestClient("http://157.26.82.44:2240/api/document/save");
            client2.Timeout = -1;
            var request2 = new RestRequest(Method.POST);
            request2.AddHeader("Authorization", "Bearer " + token);
            request2.AddHeader("Content-Type", "application/json");
            var body = jobject;
            request2.AddParameter("application/json", body, RestSharp.ParameterType.RequestBody);
            IRestResponse response2 = client2.Execute(request2);
            





        }

        static void exportPDF(String fileToTransform, String fileName, String id)
        {
            byte[] sPDFDecoded = Convert.FromBase64String(fileToTransform);

            string[] arr = fileName.Split('.');

            string newfilename = arr[0] + "_exported.pdf";


            File.WriteAllBytes("C:\\digital_corner\\Export\\" + id + newfilename, sPDFDecoded);
            Console.WriteLine(newfilename);

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
                request2.AddHeader("Authorization", "Bearer " + token);IRestResponse response2 = client2.Execute(request2);
               
                String jsonResponse2 = response2.Content.ToString();
                JObject jObject2 = Newtonsoft.Json.Linq.JObject.Parse(jsonResponse2);
                String filename = jObject2["FileName"].ToString();
                
                String fileToTransform = jObject2["File"].ToString();

                exportPDF(fileToTransform, filename, idList[i]);

                changeState(token, idList[i]);





            }

            string jsonInput = mainJson.ToString();
            

            var workbook = new Workbook();

            var worksheet = workbook.Worksheets[0];

            var layoutOptions = new JsonLayoutOptions();
            layoutOptions.ArrayAsTable = true;

            JsonUtility.ImportData(jsonInput, worksheet.Cells, 0, 0, layoutOptions);

            workbook.Save("C:\\digital_corner\\Export\\output.csv", SaveFormat.CSV);
            Console.WriteLine("csv exported");

            


        }


        

       

    }


    
}








