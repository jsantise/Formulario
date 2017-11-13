using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BL_FormContacto
{
    /*DESARROLLO
        LANDIG CARRERA                    A571F1E2-5A62-E711-80C9-000C299AB360
        LANDING CORPORATIVA Y FACULTAD    8DAAB0D8-5A62-E711-80C9-000C299AB360 */

    public class Generales
    {
        public static Entity BuscaRutPorTipoProspecto(string rut, int tipocontacto, IOrganizationService _Service)
        {

            ColumnSet col = new ColumnSet(true);
            ConditionExpression condicion = new ConditionExpression();
            condicion.AttributeName = "new_rut";
            condicion.Operator = ConditionOperator.Equal;
            condicion.Values.Add(rut);

            ConditionExpression condicion2 = new ConditionExpression();
            condicion2.AttributeName = "new_tipodecontacto";
            condicion2.Operator = ConditionOperator.Equal;
            condicion2.Values.Add(tipocontacto);

            FilterExpression filtro = new FilterExpression();
            filtro.AddCondition(condicion);
            filtro.AddCondition(condicion2);
            QueryExpression query = new QueryExpression();
            query.EntityName = "contact";
            query.Criteria = filtro;
            query.ColumnSet = col;

            EntityCollection Campos = _Service.RetrieveMultiple(query);

            if (Campos.Entities.Count > 0)
            {
                Entity contacto = new Entity("contact");
                foreach (Entity opp in Campos.Entities)
                {
                    contacto = new Entity("contact");
                    contacto = opp;
                }
                return contacto;
            }
            else
            {
                return null;
            }
        }
        public static Entity BuscaRutPorContactos(string rut, IOrganizationService _Service)
        {

            ColumnSet col = new ColumnSet(true);
            ConditionExpression condicion = new ConditionExpression();
            condicion.AttributeName = "new_rut";
            condicion.Operator = ConditionOperator.Equal;
            condicion.Values.Add(rut);

            ConditionExpression condicion2 = new ConditionExpression();
            condicion2.AttributeName = "new_tipodecontacto";
            condicion2.Operator = ConditionOperator.NotEqual;
            condicion2.Values.Add(0);

            FilterExpression filtro = new FilterExpression();
            filtro.AddCondition(condicion);
            filtro.AddCondition(condicion2);

            QueryExpression query = new QueryExpression();
            query.EntityName = "contact";
            query.Criteria = filtro;
            query.ColumnSet = col;

            EntityCollection Campos = _Service.RetrieveMultiple(query);

            if (Campos.Entities.Count > 0)
            {
                Entity contacto = new Entity("contact");
                foreach (Entity opp in Campos.Entities)
                {
                    contacto = new Entity("contact");
                    contacto = opp;
                }
                return contacto;
            }
            else
            {
                return null;
            }
        }
        public static Entity OrigenProspecto(string nombre, IOrganizationService _Service)
        {
            ColumnSet col = new ColumnSet(true);
            ConditionExpression condicion = new ConditionExpression();
            condicion.AttributeName = "cmtxucen_nombre";
            condicion.Operator = ConditionOperator.Equal;
            condicion.Values.Add(nombre);


            FilterExpression filtro = new FilterExpression();
            filtro.AddCondition(condicion);


            QueryExpression query = new QueryExpression();
            query.EntityName = "cmtxucen_leadsourcecode";
            query.Criteria = filtro;
            query.ColumnSet = col;

            EntityCollection Campos = _Service.RetrieveMultiple(query);

            if (Campos.Entities.Count > 0)
            {
                Entity origen = new Entity("cmtxucen_leadsourcecode");
                foreach (Entity opp in Campos.Entities)
                {
                    origen = new Entity("cmtxucen_leadsourcecode");
                    origen = opp;
                }
                return origen;
            }
            else
            {
                return null;
            }
        }
        public static Entity Periodo(string nombre, IOrganizationService _Service)
        {
            ColumnSet col = new ColumnSet(true);
            ConditionExpression condicion = new ConditionExpression();
            condicion.AttributeName = "new_name";
            condicion.Operator = ConditionOperator.Equal;
            condicion.Values.Add(nombre);


            FilterExpression filtro = new FilterExpression();
            filtro.AddCondition(condicion);


            QueryExpression query = new QueryExpression();
            query.EntityName = "new_periodo";
            query.Criteria = filtro;
            query.ColumnSet = col;

            EntityCollection Campos = _Service.RetrieveMultiple(query);

            if (Campos.Entities.Count > 0)
            {
                Entity origen = new Entity("new_periodo");
                foreach (Entity opp in Campos.Entities)
                {
                    origen = new Entity("new_periodo");
                    origen = opp;
                }
                return origen;
            }
            else
            {
                return null;
            }
        }
        public static EntityCollection TraerCarrera()
        {
            ServicioCRM_modulo conect = new ServicioCRM_modulo();
            IOrganizationService _Service = conect.ObtenerOrganization();
            ColumnSet col = new ColumnSet(true);

            ConditionExpression condicion2 = new ConditionExpression();
            condicion2.AttributeName = "statecode";
            condicion2.Operator = ConditionOperator.Equal;
            condicion2.Values.Add(0);

            ConditionExpression condicion4 = new ConditionExpression();
            condicion4.AttributeName = "new_niveldeformacion";
            condicion4.Operator = ConditionOperator.Equal;
            condicion4.Values.Add("1");

            FilterExpression filtro = new FilterExpression();
            filtro.AddCondition(condicion2);
            filtro.AddCondition(condicion4);

            QueryExpression query = new QueryExpression();
            query.EntityName = "new_carrera";
            query.Criteria = filtro;
            query.ColumnSet = col;
            query.AddOrder("new_name", OrderType.Ascending);


            return _Service.RetrieveMultiple(query);

        }
        public static EntityCollection TraerCarreraTecnica()
        {
            ServicioCRM_modulo conect = new ServicioCRM_modulo();
            IOrganizationService _Service = conect.ObtenerOrganization();
            ColumnSet col = new ColumnSet(true);

            ConditionExpression condicion2 = new ConditionExpression();
            condicion2.AttributeName = "statecode";
            condicion2.Operator = ConditionOperator.Equal;
            condicion2.Values.Add(0);

            ConditionExpression condicion4 = new ConditionExpression();
            condicion4.AttributeName = "new_niveldeformacion";
            condicion4.Operator = ConditionOperator.Equal;
            condicion4.Values.Add("0");

            FilterExpression filtro = new FilterExpression();
            filtro.AddCondition(condicion2);
            filtro.AddCondition(condicion4);

            QueryExpression query = new QueryExpression();
            query.EntityName = "new_carrera";
            query.Criteria = filtro;
            query.ColumnSet = col;
            query.AddOrder("new_name", OrderType.Ascending);

            return _Service.RetrieveMultiple(query);

        }
        public static EntityCollection TraerCarreraVesp()
        {

            ServicioCRM_modulo conect = new ServicioCRM_modulo();
            IOrganizationService _Service = conect.ObtenerOrganization();
            ColumnSet col = new ColumnSet(true);

            ConditionExpression condicion2 = new ConditionExpression();
            condicion2.AttributeName = "statecode";
            condicion2.Operator = ConditionOperator.Equal;
            condicion2.Values.Add(0);

            string jornadaid = "8878451A-4D44-E611-80C1-000C299AB360"; //Jornada Vespertina produccion

            ConditionExpression condicion3 = new ConditionExpression();
            condicion3.AttributeName = "new_jornadaid";
            condicion3.Operator = ConditionOperator.Equal;
            condicion3.Values.Add(jornadaid);

            //MOSTRAR SOLO LAS CARRERAS PROFESIONALES EN VESPERTINO
            /* ConditionExpression condicion4 = new ConditionExpression();
            condicion4.AttributeName = "new_niveldeformacion";
            condicion4.Operator = ConditionOperator.Equal;
            condicion4.Values.Add("1"); */

            FilterExpression filtro = new FilterExpression();
            filtro.AddCondition(condicion2);
            filtro.AddCondition(condicion3);
            //filtro.AddCondition(condicion4);

            QueryExpression query = new QueryExpression();
            query.EntityName = "new_carrera";
            query.Criteria = filtro;
            query.ColumnSet = col;
            query.AddOrder("new_name", OrderType.Ascending);

            return _Service.RetrieveMultiple(query);

        }

        public static void Contacto(string rut, string nombre, string apellidos, string email, string telefono, string carreraid, string mensaje)
        {

            try
            {
                ServicioCRM_modulo conect = new ServicioCRM_modulo();
                IOrganizationService _Service = conect.ObtenerOrganization();
                Entity Contacto = new Entity("contact");
                Entity cWeb = new Entity("new_prepostulacion");

                bool existe = false;

                //Busca en los prospectos
                Entity contacto = BuscaRutPorTipoProspecto(rut, 0, _Service);

                if (contacto == null)
                {
                    //Busca en los contactos
                    contacto = BuscaRutPorContactos(rut, _Service);

                    if (contacto == null)
                    {
                        existe = false;
                        contacto = new Entity("contact");

                    }
                    else
                    {
                        existe = true;
                    }
                }
                else
                {

                    existe = true;
                }

                if (!existe)
                {
                    if (rut != null && rut.Trim() != "")
                    {
                        Contacto.Attributes["new_rut"] = rut.ToString();
                    }

                }

                if (nombre != null && nombre.Trim() != "")
                {
                    Contacto.Attributes["firstname"] = nombre.ToString();
                }
                if (apellidos != null && apellidos.Trim() != "")
                {
                    Contacto.Attributes["lastname"] = apellidos.ToString();
                }

                if (email != null && email.Trim() != "")
                {
                    Contacto.Attributes["emailaddress1"] = email.ToString();
                }
                if (telefono != null && telefono.Trim() != "")
                {
                    Contacto.Attributes["mobilephone"] = telefono.ToString();
                }



                if (carreraid != null && carreraid.Trim() != "")
                {
                    Contacto.Attributes["cmtxucen_preferenciacarrera1id"] = new EntityReference("new_carrera", new Guid(carreraid));
                }


                //SE DECLARA LA VARIABLE PERIODO AL COMIENZO DEL METODO 

                Entity periodoid = Periodo("2018-1", _Service);
                Contacto.Attributes["cmtxucen_periodoid"] = new EntityReference(periodoid.LogicalName, periodoid.Id);


                if (mensaje != null && mensaje.Trim() != "")
                {
                    Contacto.Attributes["description"] = mensaje.ToString();
                }

                Entity prospecto = OrigenProspecto("Formulario de contacto web", _Service);
                Guid ContactoId = new Guid();
                if (!existe)
                {
                    Contacto.Attributes["cmtxucen_leadsourcecode"] = new EntityReference(prospecto.LogicalName, prospecto.Id);
                    ContactoId = _Service.Create(Contacto);

                }
                else
                {
                    Contacto.Id = contacto.Id;
                    _Service.Update(Contacto);
                    ContactoId = Contacto.Id;
                }


                #region CREAR PREPOSTULACIÓN (CONSULTA WEB)

                if (!existe)
                {
                    if (rut != null && rut.Trim() != "")
                    {
                        cWeb.Attributes["new_rut"] = rut.ToString();
                    }

                    if (nombre != null && nombre.Trim() != "")
                    {
                        cWeb.Attributes["new_nombres"] = nombre.ToString();
                    }
                    if (apellidos != null && apellidos.Trim() != "")
                    {
                        cWeb.Attributes["new_apellidopaterno"] = apellidos.ToString();
                    }

                    if (email != null && email.Trim() != "")
                    {
                        cWeb.Attributes["new_correoelectronico"] = email.ToString();
                    }
                    if (telefono != null && telefono.Trim() != "")
                    {
                        cWeb.Attributes["new_telefonomovil"] = telefono.ToString();
                    }

                    if (carreraid != null && carreraid.Trim() != "")
                    {
                        cWeb.Attributes["new_carrera1"] = new EntityReference("new_carrera", new Guid(carreraid));
                    }

                    if (mensaje != null && mensaje.Trim() != "")
                    {
                        cWeb.Attributes["new_consulta"] = mensaje.ToString();
                    }

                    if (!String.IsNullOrEmpty(nombre) && !String.IsNullOrEmpty(apellidos))
                    {
                        cWeb.Attributes["subject"] = "Prospecto WEB " + nombre.ToString() + " " + apellidos.ToString();
                    }

                    cWeb.Attributes["regardingobjectid"] = new EntityReference(Contacto.LogicalName, ContactoId);

                    _Service.Create(cWeb);

                    //CREAR ACTIVIDAD DE IMPACTO
                    Entity actim = new Entity("cmtxucen_actividadimpacto");

                    actim.Attributes["regardingobjectid"] = new EntityReference(Contacto.LogicalName, ContactoId);
                    actim.Attributes["cmtxucen_actividaddeimpacto"] = new EntityReference(prospecto.LogicalName, prospecto.Id);
                    actim.Attributes["subject"] = "Prospecto WEB " + nombre.ToString() + " " + apellidos.ToString();
                    actim.Attributes["statecode"] = 1;
                    actim.Attributes["statecode"] = 2;

                    _Service.Create(actim);
                }
                else
                {

                    if (rut != null && rut.Trim() != "")
                    {
                        cWeb.Attributes["new_rut"] = rut.ToString();
                    }

                    if (nombre != null && nombre.Trim() != "")
                    {
                        cWeb.Attributes["new_nombres"] = nombre.ToString();
                    }
                    if (apellidos != null && apellidos.Trim() != "")
                    {
                        cWeb.Attributes["new_apellidopaterno"] = apellidos.ToString();
                    }

                    if (email != null && email.Trim() != "")
                    {
                        cWeb.Attributes["new_correoelectronico"] = email.ToString();
                    }
                    if (telefono != null && telefono.Trim() != "")
                    {
                        cWeb.Attributes["new_telefonomovil"] = telefono.ToString();
                    }

                    if (carreraid != null && carreraid.Trim() != "")
                    {
                        cWeb.Attributes["new_carrera1"] = new EntityReference("new_carrera", new Guid(carreraid));
                    }

                    if (mensaje != null && mensaje.Trim() != "")
                    {
                        cWeb.Attributes["new_consulta"] = mensaje.ToString();
                    }

                    if (!String.IsNullOrEmpty(nombre) && !String.IsNullOrEmpty(apellidos))
                    {
                        cWeb.Attributes["subject"] = "Prospecto WEB " + nombre.ToString() + " " + apellidos.ToString();
                    }

                    cWeb.Attributes["regardingobjectid"] = new EntityReference(Contacto.LogicalName, ContactoId);

                    _Service.Create(cWeb);

                    //CREAR ACTIVIDAD DE IMPACTO
                    Entity actim = new Entity("cmtxucen_actividadimpacto");

                    actim.Attributes["regardingobjectid"] = new EntityReference(Contacto.LogicalName, ContactoId);
                    actim.Attributes["cmtxucen_actividaddeimpacto"] = new EntityReference(prospecto.LogicalName, prospecto.Id);
                    actim.Attributes["subject"] = "Prospecto WEB " + nombre.ToString() + " " + apellidos.ToString();
                    actim.Attributes["statecode"] = 1;
                    actim.Attributes["statecode"] = 2;

                    _Service.Create(actim);

                }
                #endregion


            }
            catch (Exception ex)
            {

            }

        }
    }
}
