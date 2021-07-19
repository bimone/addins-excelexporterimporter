using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

#pragma warning disable

namespace ExcelExporterImporter.Common
{
    [GeneratedCode("System.Xml", "4.0.30319.34230")]
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false)]
    public class ParametersSettings
    {
        private static XmlSerializer serializer;
        private List<IgnoredParameters> ignoredParameters;
        private List<ParameterTranslation> parametersTranslations;
        private List<ReadonlyParameter> readonlyParameters;

        public ParametersSettings()
        {
            ignoredParameters = new List<IgnoredParameters>();
            readonlyParameters = new List<ReadonlyParameter>();
            parametersTranslations = new List<ParameterTranslation>();
        }

        [XmlArrayItem("ParameterTranslation", IsNullable = false)]
        public List<ParameterTranslation> ParametersTranslations
        {
            get => parametersTranslations;
            set => parametersTranslations = value;
        }

        [XmlElement("IgnoredParameters")]
        public List<IgnoredParameters> IgnoredParameters
        {
            get => ignoredParameters;
            set => ignoredParameters = value;
        }

        [XmlElement("ReadonlyParameter")]
        public List<ReadonlyParameter> ReadonlyParameters
        {
            get => readonlyParameters;
            set => readonlyParameters = value;
        }


        private static XmlSerializer Serializer
        {
            get
            {
                if (serializer == null)
                    serializer = new XmlSerializerFactory().CreateSerializer(typeof(ParametersSettings));
                return serializer;
            }
        }

        /// <summary>
        ///     Deserializes xml markup from file into an ParametersSettings object
        /// </summary>
        /// <param name="fileName">string xml file to load and deserialize</param>
        /// <param name="obj">Output ParametersSettings object</param>
        /// <param name="exception">output Exception value if deserialize failed</param>
        /// <returns>true if this Serializer can deserialize the object; otherwise, false</returns>
        private static bool LoadFromFile(string fileName, out ParametersSettings obj, out Exception exception)
        {
            exception = null;
            obj = default;
            try
            {
                obj = LoadFromFile(fileName);
                return true;
            }
            catch (Exception ex)
            {
                exception = ex;
                return false;
            }
        }

        /// <summary>
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static bool LoadFromFile(string fileName, out ParametersSettings obj)
        {
            Exception exception = null;
            return LoadFromFile(fileName, out obj, out exception);
        }

        /// <summary>
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static ParametersSettings LoadFromFile(string fileName)
        {
            FileStream file = null;
            StreamReader sr = null;
            try
            {
                file = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                sr = new StreamReader(file);
                var xmlString = sr.ReadToEnd();
                sr.Close();
                file.Close();
                return Deserialize(xmlString);
            }
            finally
            {
                if (file != null) file.Dispose();
                if (sr != null) sr.Dispose();
            }
        }

        #region Serialize/Deserialize

        /// <summary>
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        private static ParametersSettings Deserialize(string input)
        {
            StringReader stringReader = null;
            try
            {
                stringReader = new StringReader(input);
                return (ParametersSettings) Serializer.Deserialize(XmlReader.Create(stringReader));
            }
            finally
            {
                if (stringReader != null) stringReader.Dispose();
            }
        }

        #endregion
    }

    [GeneratedCode("System.Xml", "4.0.30319.34230")]
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    public class ParameterTranslation
    {
        private string _location;
        private string _name;
        private List<ParametersSettingsParameterTranslations> _translations;

        public ParameterTranslation()
        {
            _translations = new List<ParametersSettingsParameterTranslations>();
        }

        [XmlElement("Translations")]
        public List<ParametersSettingsParameterTranslations> Translations
        {
            get => _translations;
            set => _translations = value;
        }

        [XmlAttribute]
        public string Name
        {
            get => _name;
            set => _name = value;
        }

        [XmlAttribute]
        public string Location
        {
            get => _location;
            set => _location = value;
        }
    }

    [GeneratedCode("System.Xml", "4.0.30319.34230")]
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    public class ParametersSettingsParameterTranslations
    {
        private string _text;
        private string _value;

        [XmlAttribute]
        public string Text
        {
            get => _text;
            set => _text = value;
        }

        [XmlAttribute]
        public string Value
        {
            get => _value;
            set => _value = value;
        }
    }

    [GeneratedCode("System.Xml", "4.0.30319.34230")]
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    public class IgnoredParameters
    {
        private string _location;
        private List<ParametersSettingsIgnoredParametersParameter> _parameter;

        public IgnoredParameters()
        {
            _parameter = new List<ParametersSettingsIgnoredParametersParameter>();
        }

        [XmlElement("Parameter")]
        public List<ParametersSettingsIgnoredParametersParameter> Parameter
        {
            get => _parameter;
            set => _parameter = value;
        }

        [XmlAttribute]
        public string Location
        {
            get => _location;
            set => _location = value;
        }
    }

    [GeneratedCode("System.Xml", "4.0.30319.34230")]
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    public class ReadonlyParameter
    {
        private int _id;

        [XmlAttribute]
        public int Id
        {
            get => _id;
            set => _id = value;
        }
    }

    [GeneratedCode("System.Xml", "4.0.30319.34230")]
    [Serializable]
    [DebuggerStepThrough]
    [DesignerCategory("code")]
    [XmlType(AnonymousType = true)]
    public class ParametersSettingsIgnoredParametersParameter
    {
        private string _name;

        [XmlAttribute]
        public string Name
        {
            get => _name;
            set => _name = value;
        }
    }
}

#pragma warning restore