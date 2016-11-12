namespace AzureAddin.Model
{
    class ModelData
    {

        private string nameofstorage;
        private string primarykey;
        private string nameofcontainer;
        private string filename;

        public string Nameofstorage
        {
            get
            {
                return nameofstorage;
            }

            set
            {
                nameofstorage = value;
            }
        }

        public string Primarykey
        {
            get
            {
                return primarykey;
            }

            set
            {
                primarykey = value;
            }
        }

        public string Nameofcontainer
        {
            get
            {
                return nameofcontainer;
            }

            set
            {
                nameofcontainer = value;
            }
        }

        public string Filename
        {
            get
            {
                return filename;
            }

            set
            {
                filename = value;
            }
        }
    }
}
