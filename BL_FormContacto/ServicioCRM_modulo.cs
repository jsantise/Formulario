using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Configuration;
using System.ServiceModel.Description;

namespace BL_FormContacto
{
    public class ServicioCRM_modulo
    {
        #region "ServicioCRM"
        private static OrganizationService _ServicioCRM;
        //Private _MetadataCRM As MetadataService.MetadataService
        public static OrganizationService ServicioCRM
        {
            get
            {
                if ((_ServicioCRM == null))
                    crea_ServicioCRM();
                return _ServicioCRM;
            }

        }
        public static void crea_ServicioCRM()
        {
            try
            {

                //Conectar a CRM: Office 365 - Windows Live ID - CRM Online - CRM on Premise
                //Windows integrated security - (IFD) with claims             
                String connectionString = "Url=https://srm.ucentral.cl:444; Username=ucentral.cl\\crm_admin; Password=cmetrix.2016;";

                CrmConnection connection = CrmConnection.Parse(connectionString);
                OrganizationService _serviceProxy = new OrganizationService(connection);

                _ServicioCRM = ((OrganizationService)_serviceProxy);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        public IOrganizationService ObtenerOrganization()
        {
            string CredencialUsuario = ConfigurationManager.AppSettings["CredencialUsuario"];
            string CredencialContrasena = ConfigurationManager.AppSettings["CredencialContrasena"];
            string CredencialDominio = ConfigurationManager.AppSettings["CredencialDominio"];
            string CRMOrganizationName = ConfigurationManager.AppSettings["CRMOrganizationName"];
            string CRMServidor = ConfigurationManager.AppSettings["CRMServidor"];
            string UsaSSL = ConfigurationManager.AppSettings["UsaSSL"];
            string TipoAutenticacion = ConfigurationManager.AppSettings["TipoAutenticacion"];
            string http = ConfigurationManager.AppSettings["TipoLogueo"];
            string TipoCRM = ConfigurationManager.AppSettings["TipoCRM"];
            string CertificateValidation = ConfigurationManager.AppSettings["CertificateValidation"];
            string TimeOut = ConfigurationManager.AppSettings["TimeOut_min"];


            ClientCredentials Credencial = new ClientCredentials();
            Credencial.UserName.UserName = CredencialUsuario;
            Credencial.UserName.Password = CredencialContrasena;

            Uri HomeRealmUri;
            HomeRealmUri = null;

            Uri OrganizationUri = new Uri("https://" + CRMServidor + "/" + CRMOrganizationName + "/XRMServices/2011/Organization.svc");

            OrganizationServiceProxy _serviceProxy;
            _serviceProxy = new OrganizationServiceProxy(OrganizationUri, HomeRealmUri, Credencial, null);

            _serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());
            return ((IOrganizationService)_serviceProxy);


        }

    }
}