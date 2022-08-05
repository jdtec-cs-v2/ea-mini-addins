namespace MiniAddins.Models
{
    public class ToolCat
    {
        private string icon;
        private string displayName;
        private string displayNameResourceKey;
        private string key;
        private string view;

        public string Icon
        {
            get { return icon; }
            set { icon = value; }
        }

        public string DisplayName
        {
            get { return displayName; }
            set { displayName = value; }
        }

        public string DisplayNameResourceKey
        {
            get { return displayNameResourceKey; }
            set { displayNameResourceKey = value; }
        }


        public string Key
        {
            get { return key; }
            set { key = value; }
        }

        public string View
        {
            get { return view; }
            set { view = value; }
        }
    }


}
