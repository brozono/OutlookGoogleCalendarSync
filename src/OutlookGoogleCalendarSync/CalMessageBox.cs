using log4net;
using System;
using System.Runtime.Serialization;
using System.Windows.Forms;

namespace OutlookGoogleCalendarSync {

    [DataContract]
    public class CalMessageBox {
        private static CalMessageBox instance;

        public static CalMessageBox Instance {
            get {
                if (instance == null) instance = new CalMessageBox();
                return instance;
            }
            set {
                instance = value;
            }
        }

        public Boolean ShowTrue(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, System.Windows.Forms.DialogResult bypass) {
            if (Settings.Instance.EnableAutoRetry) {
                return true;
            }

            return MessageBox.Show(text, caption, buttons, icon) == bypass;
        }

        public Boolean ShowFalse(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon, System.Windows.Forms.DialogResult bypass) {
            if (Settings.Instance.EnableAutoRetry) {
                return false;
            }

            return MessageBox.Show(text, caption, buttons, icon) != bypass;
        }
    }
}
