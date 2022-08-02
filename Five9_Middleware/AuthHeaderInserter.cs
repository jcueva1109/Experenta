using System;
using System.ServiceModel.Channels;
using System.ServiceModel.Dispatcher;
using System.Text;

namespace Five9_Middleware.Helpers
{
    public class AuthHeaderInserter : IClientMessageInspector
    {

        public string Username
        {
            get => _username;
            set
            {
                _username = value;
                BuildAuthString();
            }
        }

        public string Password
        {

            get => _password;
            set {
                _password = value;
                BuildAuthString();
            }

        }

        public string AuthHeader => _auth;
        private string _username;
        private string _password;
        private string _auth;
        private void BuildAuthString()
        {

            _auth = string.Concat("Basic ", Convert.ToBase64String(
                Encoding.UTF8.GetBytes(string.Concat(Username, ":", Password))));

        }

        public void AfterReceiveReply(ref Message reply, object correlationState)
        {
            //Free HttpRequest
        }

        public object BeforeSendRequest(ref Message req, System.ServiceModel.IClientChannel ch)
        {

            HttpRequestMessageProperty httpRequest;
            object o;

            if (req.Properties.TryGetValue(HttpRequestMessageProperty.Name, out o))
            {
                httpRequest = o as HttpRequestMessageProperty;
                if (httpRequest != null) httpRequest.Headers["Authorization"] = AuthHeader;
            }
            else
            {
                httpRequest = new HttpRequestMessageProperty();
                httpRequest.Headers["Authorization"] = AuthHeader;
                req.Properties.Add(HttpRequestMessageProperty.Name, httpRequest);
            
            }
            
            return httpRequest;

        }

    }//EOF
}
