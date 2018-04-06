using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace SpiraExcelAddIn
{
    /// <summary>
    /// Stores all the settings used in the application. Cannot use .NET settings because VSTO does not have access
    /// to the app.config file
    /// </summary>
    public class Configuration
    {
        private const string SETTINGS_FILE = "SpiraExcelAddIn2007.config";

        static Configuration _Default;

        /// <summary>
        /// Static constructor
        /// </summary>
        static Configuration()
        {
            _Default = new Configuration();
            _Default.Load();
        }

        /// <summary>
        /// Constructor
        /// </summary>
        private Configuration()
        {
            this.SpiraUrl = "http://localhost/SpiraTest";
            this.SpiraUserName = "";
            this.SpiraPassword = "";
            this.StripRichText = false;
        }

        /// <summary>
        /// Returns the static instance of the config class
        /// </summary>
        public static Configuration Default
        {
            get
            {
                return _Default;
            }
        }

        /// <summary>
        /// Saves the setting values
        /// </summary>
        public void Save()
        {
            //See if we have the file in the user's profile
            string localUserFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string inflectraFolder = Path.Combine(localUserFolder, "Inflectra");
            if (!Directory.Exists(inflectraFolder))
            {
                Directory.CreateDirectory(inflectraFolder);
            }

            //We need to serialize the current object to the file stream
            string filePath = Path.Combine(inflectraFolder, SETTINGS_FILE);
            FileStream stream = new FileStream(filePath, FileMode.Create);
            System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(this.GetType());
            serializer.Serialize(stream, this);
            stream.Close();
        }

        /// <summary>
        /// Loads the setting values (if the setting file exists otherwise does nothing)
        /// </summary>
        public void Load()
        {
            //See if we have the file in the user's profile
            string localUserFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string inflectraFolder = Path.Combine(localUserFolder, "Inflectra");
            if (Directory.Exists(inflectraFolder))
            {
                //We need to deserialize from the file stream to the current object
                string filePath = Path.Combine(inflectraFolder, SETTINGS_FILE);
                if (File.Exists(filePath))
                {
                    FileStream stream = new FileStream(filePath, FileMode.Open);
                    System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(this.GetType());
                    Configuration configuration = (Configuration)serializer.Deserialize(stream);
                    this.SpiraUrl = configuration.SpiraUrl;
                    this.SpiraUserName = configuration.SpiraUserName;
                    this.SpiraPassword = configuration.SpiraPassword;
                    this.StripRichText = configuration.StripRichText;
                    this.TestRunDate = configuration.TestRunDate;
                    stream.Close();
                }
            }
        }

        public string SpiraUrl
        {
            get;
            set;
        }

        public string SpiraUserName
        {
            get;
            set;
        }
        public string SpiraPassword
        {
            get;
            set;
        }
        public bool StripRichText
        {
            get;
            set;
        }
        public DateTime? TestRunDate
        {
            get;
            set;
        }
    }
}
