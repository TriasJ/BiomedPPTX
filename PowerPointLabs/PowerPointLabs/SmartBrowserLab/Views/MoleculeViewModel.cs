using System.ComponentModel;

namespace PowerPointLabs.SmartBrowserLab.Views
{
    public class MoleculeViewModel : INotifyPropertyChanged
    {
        private int _cid;
        private string _name;
        private string _formula;
        private string _weight;
        private string _iupacName;
        private string _thumbnailPath;
        private bool _hasConformer3D;

        public event PropertyChangedEventHandler PropertyChanged;

        public int Cid
        {
            get { return _cid; }
            set
            {
                _cid = value;
                OnPropertyChanged("Cid");
            }
        }

        public string Name
        {
            get { return _name; }
            set
            {
                _name = value;
                OnPropertyChanged("Name");
            }
        }

        public string Formula
        {
            get { return _formula; }
            set
            {
                _formula = value;
                OnPropertyChanged("Formula");
            }
        }

        public string Weight
        {
            get { return _weight; }
            set
            {
                _weight = value;
                OnPropertyChanged("Weight");
            }
        }

        public string IupacName
        {
            get { return _iupacName; }
            set
            {
                _iupacName = value;
                OnPropertyChanged("IupacName");
            }
        }

        public string ThumbnailPath
        {
            get { return _thumbnailPath; }
            set
            {
                _thumbnailPath = value;
                OnPropertyChanged("ThumbnailPath");
            }
        }

        public bool HasConformer3D
        {
            get { return _hasConformer3D; }
            set
            {
                _hasConformer3D = value;
                OnPropertyChanged("HasConformer3D");
            }
        }

        public string DisplayName
        {
            get
            {
                if (!string.IsNullOrEmpty(_name))
                {
                    return _name;
                }

                return string.Format("CID {0}", _cid);
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
