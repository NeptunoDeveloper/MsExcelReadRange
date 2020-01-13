using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MsExcelAutomatizacion
{
    public abstract class FileExcelTrace
    {
        private string _message;
        private string _messageDetail;
        private FileExcelLevel _messageLevel = FileExcelLevel.None;

        public string getMessage() { return _message;  }
        public string getMessageDetail() { return _messageDetail; }
        public string getLevel()
        {

            if (_messageLevel == FileExcelLevel.Error)
                return "ERROR";
            else
                return "OK";
        }
        public void setOk(string pStrMessage, string pStrMessageDetail)
        {
            _message = pStrMessage;
            _messageDetail = pStrMessageDetail;
            _messageLevel = FileExcelLevel.Ok;
        }

        public void setOk(string pStrMessage)
        {
            _message = pStrMessage;
            _messageDetail = string.Empty;
            _messageLevel = FileExcelLevel.Ok;
        }

        public void setError(string pStrMessage, string pStrMessageDetail)
        {
            _message = pStrMessage;
            _messageDetail = pStrMessageDetail;
            _messageLevel = FileExcelLevel.Error;
        }

        public void setError(string pStrMessage)
        {
            _message = pStrMessage;
            _messageDetail = string.Empty;
            _messageLevel = FileExcelLevel.Error;
        }

    }


    public enum FileExcelLevel
    {
        None = 0,
        Ok = 1,
        Error = 2
    }
}
