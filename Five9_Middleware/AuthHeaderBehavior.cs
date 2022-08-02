using System.Diagnostics;
using System.ServiceModel.Channels;
using System.ServiceModel.Description;
using System.ServiceModel.Dispatcher;

namespace Five9_Middleware.Helpers
{
    public class AuthHeaderBehavior : IEndpointBehavior
    {

        private readonly AuthHeaderInserter _authHeaderInserter;

        public AuthHeaderBehavior(AuthHeaderInserter headerInserter)
        {
            _authHeaderInserter = headerInserter;
        }

        public void AddBindingParameters(ServiceEndpoint eP, BindingParameterCollection bP)
        {
            //
        }

        public void ApplyClientBehavior(ServiceEndpoint eP, ClientRuntime clientRuntime)
        {

            clientRuntime.ClientMessageInspectors.Add(_authHeaderInserter);
            Debug.WriteLine("AuthHeaderInserter plus MessageInspector for: " + eP.Address);

        }

        public void ApplyDispatchBehavior(ServiceEndpoint eP, EndpointDispatcher ePDispatcher)
        {
            //
        }

        public void Validate(ServiceEndpoint eP)
        {
            //
        }

    }//EOF
}
