using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.WindowsAzure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.WindowsAzure.Storage.Blob;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using AzureAddin.Model;
using AzureAddin.Commands;

namespace AzureAddin.Viewmodel
{
    class ViewDatamodel : INotifyPropertyChanged
    {

        public event PropertyChangedEventHandler PropertyChanged;

        public void RaisePropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        private RelayCommand upload;
        public RelayCommand Upload
        {
            get
            {
                return upload ??
                  (upload = new RelayCommand(obj =>UploadtoAzure(Nameofstorage, Primarykey, Nameofcontainer, Filename)));
            }
        }

        private ModelData model;

        public ViewDatamodel()
        {
            model = new ModelData();
        }

        //Creating new model(viewmodel) inside Viewmodel for work with View 
        public string Nameofstorage
        {
            get
            {
                return model.Nameofstorage;
            }

            set
            {
                model.Nameofstorage = value;
                RaisePropertyChanged("Nameofstorage");
            }
        }

        public string Primarykey
        {
            get
            {
                return model.Primarykey;
            }

            set
            {
                model.Primarykey = value;
                RaisePropertyChanged("Nameofstorage");
            }
        }

        public string Nameofcontainer
        {
            get
            {
                return model.Nameofcontainer;
            }

            set
            {
                model.Nameofcontainer = value;
                RaisePropertyChanged("Nameofcontainer");
            }
        }

        public string Filename
        {
            get
            {
                return model.Filename;
            }

            set
            {
                model.Filename = value;
                RaisePropertyChanged("Filename");
            }
        }

        //Method for upload docfile in storage
        private async void UploadtoAzure(string storagename, string primkey, string containername, string filename)
        {
            try
            {
                //Create container on datastorage for save document 
                StorageCredentials sc = new StorageCredentials(storagename, primkey);
                CloudStorageAccount stroracc = new CloudStorageAccount(sc, true);
                CloudBlobClient blobclient = stroracc.CreateCloudBlobClient();
                CloudBlobContainer container = blobclient.GetContainerReference(containername);
                await container.CreateIfNotExistsAsync();

                //create way for save file on desktop
                string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                filename += ".doc";
                string file = desktopFolder + @"\" + filename;

                //Save document and close connection
                Globals.ThisAddIn.Application.ActiveDocument.SaveAs2(file, Word.WdSaveFormat.wdFormatDocument);
                Globals.ThisAddIn.Application.ActiveDocument.Close();

                //Send our document to Azure storage
                CloudBlockBlob blockblob = container.GetBlockBlobReference(filename);
                await blockblob.UploadFromFileAsync(file);
            }

            catch (StorageException)
            {
                System.Windows.MessageBox.Show("You wrote wrong Storage name or Primary Key");
            }
        }

    }
}
