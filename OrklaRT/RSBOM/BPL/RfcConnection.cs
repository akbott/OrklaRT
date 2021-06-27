using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using ERPConnect;

namespace BPL
{
    public class RfcConnection 
    {
        public static bool CheckBHPConnection()
        {
            bool ret = false;
            using (var entities = new DAL.SAPExlEntities())
            {
                var currentUser = entities.vwCurrentUser.SingleOrDefault();
                if (currentUser.BwHana != null)
                {
                    try
                    {
                        R3Connection con = GetBHPConnection(String.Empty);                      
                        con.Ping();
                        ret = true;
                    }
                    catch { }
                }
            }
            return ret;
        }
        public static bool CheckR3PConnection()
        {
            bool ret = false;
            using (var entities = new DAL.SAPExlEntities())
            {
                var currentUser = entities.vwCurrentUser.SingleOrDefault();
                if (currentUser.SAPSystem != null)
                {
                    try
                    {
                        RfcDestination rfcDestnitation = RfcDestinationManager.GetDestination(GetParameters(currentUser.SAPSystem,"2"));
                        rfcDestnitation.Ping();                        
                        ret = true;
                    }
                    catch { Exception ex; }
                }
            }
            return ret;
        }
        public static RfcDestination GetDestination(string useSapGui)
        {
            using (var entities = new DAL.SAPExlEntities())
            {
                var currentUser = entities.vwCurrentUser.SingleOrDefault();
                return RfcDestinationManager.GetDestination(GetParameters(currentUser.SAPSystem,useSapGui));
            }
        }        
        public static RfcConfigParameters GetParameters(string name,string useSapGui)
        {
            RfcConfigParameters parameters = new RfcConfigParameters();
            using (var entities = new DAL.SAPExlEntities())
            {
                var userSAPSystem = entities.vwUserSAPSystems.SingleOrDefault(c => c.SAPSystem == name);
                var rfcParameters = entities.RfcConnection.SingleOrDefault(p=> p.Name == name);
                
                parameters[RfcConfigParameters.Name] = rfcParameters.Name;
                parameters[RfcConfigParameters.User] = userSAPSystem.SAPUserName;
                //parameters[RfcConfigParameters.Password] = DecryptString(currentUser.Password);
                parameters[RfcConfigParameters.Password] = userSAPSystem.SAPPassword;
                parameters[RfcConfigParameters.Client] = rfcParameters.Client;
                parameters[RfcConfigParameters.Language] = rfcParameters.Language;
                parameters[RfcConfigParameters.PoolSize] = "10";
                parameters[RfcConfigParameters.ConnectionIdleTimeout] = "600";
                parameters[RfcConfigParameters.AppServerHost] = rfcParameters.HostServer;
                parameters[RfcConfigParameters.SystemNumber] = (rfcParameters.SysNr.Equals(0) ? "00" : Convert.ToString(rfcParameters.SysNr));
                parameters[RfcConfigParameters.UseSAPGui] = useSapGui;
                //parameters[RfcConfigParameters.PeakConnectionsLimit] = rfcParameters.UseSAPGui;
            }
            return parameters;
        }
        public static R3Connection GetBHPConnection(string lang)
        {
            R3Connection con = new R3Connection();
            using (var entities = new DAL.SAPExlEntities())
            {
                var currentUser = entities.vwCurrentUser.SingleOrDefault();
                var userSapInfo = entities.vwUserSAPSystems.SingleOrDefault(c => c.SAPSystem == currentUser.BwHana);
                var r3Con = entities.RfcConnection.SingleOrDefault(p => p.Name == userSapInfo.SAPSystem);                
                con.UserName = userSapInfo.SAPUserName;
                con.Password = userSapInfo.SAPPassword;
                con.Language = lang.Equals(String.Empty) ? r3Con.Language : lang;
                con.Client = r3Con.Client;
                con.Host = r3Con.HostServer;
                con.SystemNumber = r3Con.SysNr;                
            }
            return con;
        }

        public static R3Connection GetR3PConnection()
        {
            R3Connection con = new R3Connection();
            using (var entities = new DAL.SAPExlEntities())
            {
                var currentUser = entities.vwCurrentUser.SingleOrDefault();
                var userSapInfo = entities.vwUserSAPSystems.SingleOrDefault(c => c.SAPSystem == currentUser.SAPSystem);
                var r3Con = entities.RfcConnection.SingleOrDefault(p => p.Name == userSapInfo.SAPSystem);
                con.Protocol = ClientProtocol.NWRFC;
                con.UserName = userSapInfo.SAPUserName;
                con.Password = userSapInfo.SAPPassword;
                con.Language = r3Con.Language;
                con.Client = r3Con.Client;
                con.Host = r3Con.HostServer;
                con.SystemNumber = r3Con.SysNr;
            }
            return con;
        }
        public static string EncryptString(string password)
        {
            byte[] Results;
            System.Text.UTF8Encoding UTF8 = new System.Text.UTF8Encoding();

            MD5CryptoServiceProvider HashProvider = new MD5CryptoServiceProvider();
            byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes("Orkla"));

            TripleDESCryptoServiceProvider TDESAlgorithm = new TripleDESCryptoServiceProvider();

            TDESAlgorithm.Key = TDESKey;
            TDESAlgorithm.Mode = CipherMode.CBC;
            TDESAlgorithm.Padding = PaddingMode.PKCS7;

            byte[] DataToEncrypt = UTF8.GetBytes(password);

            try
            {
                ICryptoTransform Encryptor = TDESAlgorithm.CreateEncryptor();
                Results = Encryptor.TransformFinalBlock(DataToEncrypt, 0, DataToEncrypt.Length);
            }
            finally
            {
                TDESAlgorithm.Clear();
                HashProvider.Clear();
            }

            return Convert.ToBase64String(Results);
        }

        public static string DecryptString(string password)
        {
            byte[] Results;

            System.Text.UTF8Encoding UTF8 = new System.Text.UTF8Encoding();
            MD5CryptoServiceProvider HashProvider = new MD5CryptoServiceProvider();

            byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes("Orkla"));
            TripleDESCryptoServiceProvider TDESAlgorithm = new TripleDESCryptoServiceProvider();

            TDESAlgorithm.Key = TDESKey;
            TDESAlgorithm.Mode = CipherMode.CBC;
            TDESAlgorithm.Padding = PaddingMode.PKCS7;
            byte[] DataToDecrypt = Convert.FromBase64String(password);
            try
            {
                ICryptoTransform Decryptor = TDESAlgorithm.CreateDecryptor();
                Results = Decryptor.TransformFinalBlock(DataToDecrypt, 0, DataToDecrypt.Length);
            }
            finally
            {
                TDESAlgorithm.Clear();
                HashProvider.Clear();
            }
            return UTF8.GetString(Results);
        }
    }
}
