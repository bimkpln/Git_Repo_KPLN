using KPLN_Library_SQLiteWorker.Core.SQLiteData;
using System.ComponentModel;

namespace KPLN_Tools.Forms.Models
{
    public class SendMsgToBitrix_UserEntity : VM_UserEntity, INotifyPropertyChanged
    {
        private bool _isSelected = false;

        public SendMsgToBitrix_UserEntity(DBUser dbUser, DBSubDepartment dbSubDepartment) : base(dbUser, dbSubDepartment)
        { }

        /// <summary>
        /// Для окна wpf - пометка, что элемент выбран
        /// </summary>
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }
    }
}
