using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.AccessControl;
using System.IO;
using System.Security.Cryptography;

namespace NEA_Prototype
{
    public class CreateUser
    {
        public string e_username;   //Username
        public string e_password;   //Password
        public string e_firstName;  //First Name
        public string e_lastName;   //Last Name
        public string e_role;   //Role
        public string userFilePath = Environment.CurrentDirectory;    //The path of where user data files are stored
        public bool isValid = true; //Boolean that returns whether a valid property was entered
        public bool validLogin = false; //Bool that returns whether a valid login was entered
        public bool isEncrypting;   //Bool set to true when encyrpting and false when decrypting
        public string decryptedFile;    //Plaintext file in string
        private string AesIV = @"5463832229458937"; //Encryption vector
        private string AesKey = @"8643530582374932";    //Encryption key
        public FileNotFoundException flEx;  //File not found exception 

        //Checks username input length and validity
        public void CheckUsername(string username)
        {
            isValid = true;
            if (username.Length >= 3 & username.Length <= 20)
                e_username = username;
            else
                isValid = false;
        }

        //Checks password input length and validity
        public void CheckPassword(string password)
        {
            isValid = true;
            if (password.Length >= 6 & password.Length <= 25)
                if (password.Any(char.IsDigit) & password.Any(char.IsUpper))
                    e_password = password;
                else
                    isValid = false;
            else
                isValid = false;
        }
        
        //Checks firstname input length and validity
        public void CheckFirstName(string firstName)
        {
            isValid = true;
            if (firstName.Length >= 2 & firstName.Length <= 25)
                if (firstName.Any(char.IsDigit) == false)
                    e_firstName = firstName;
                else
                    isValid = false;
            else
                isValid = false;    
        }

        //Checks lastname input length and validity
        public void CheckLastName(string lastName)
        {
            isValid = true;
            if (lastName.Length >= 2 & lastName.Length <= 25)
                if (lastName.Any(char.IsDigit) != true)
                    e_lastName = lastName;
                else
                    isValid = false;
            else
                isValid = false;
        }

        //Checks role input length and validity
        public void CheckRole(string role)
        {
            isValid = true;
            if (role.Length >= 2 & role.Length <= 25)
                if (role.Any(char.IsDigit) != true)
                    e_role = role;
                else
                    isValid = false;
            else
                isValid = false;
        }

        public void ReadWriteUser(string userName, string password)
        {              
            //Checks whether we are encrypting or decrypting
            if (isEncrypting == true)
            {
                //creates a string to encrypt to file
                string createdUser = (e_username + "," + e_password + "," + e_firstName + "," + e_lastName + "," + e_role);
                //Open a file with the same name as the username, if one doesnt exist create one and set file stream to write
                using (FileStream fs = File.Open((userFilePath + @"\" + e_username + ".txt") , FileMode.Append, FileAccess.Write))
                {
                    using (Aes aes = Aes.Create())
                    {
                        //Aes handles encyprtion through an array of bits
                        string encrypted = EncryptString(createdUser);
                        byte[] encryptedBytes = new UTF8Encoding(true).GetBytes(encrypted);
                        //ciphertext is written to file
                        fs.Write(encryptedBytes, 0, encryptedBytes.Length);
                    }
                }
            }
            else
            {
                try
                {
                    //Open a file with the same name as the username and set file stream to read write
                    using (FileStream fs = File.Open((userFilePath + @"\" + userName + ".txt"), FileMode.Open, FileAccess.ReadWrite))
                    {
                        using (Aes aes = Aes.Create())
                        {

                            string fileContents;

                            using (StreamReader reader = new StreamReader(fs))
                            {
                                fileContents = reader.ReadToEnd();  //contains ciphertext from file
                                byte[] CipherText = new UTF8Encoding(true).GetBytes(fileContents);  //converts to an array of bites
                                decryptedFile = DecryptUsingCBC(fileContents);
                                List<String> userLine = decryptedFile.Split(',').ToList();   //splits the string into elements
                                //assigns elements in list to our class properties
                                e_username = userLine[0];
                                e_password = userLine[1];
                                e_firstName = userLine[2];
                                e_lastName = userLine[3];
                                e_role = userLine[4];
                                //Checks whether the files username and password matches with the enetered credentials
                                if (userLine[0] == userName & userLine[1] == password)
                                    validLogin = true;
                                else
                                    validLogin = false;
                            }
                        }
                    }
                } catch (System.IO.FileNotFoundException ex)
                {
                    flEx = ex;
                    //If file is not found that means the user login does not exist and tells the end user
                    isValid = false;
                }
                                   
            }
        }

        //AES method to encrypt
        public byte[] EncryptBytes(string toEncrypt)
        {
            //Gets bytes of string to encrypt
            byte[] src = Encoding.UTF8.GetBytes(toEncrypt);
            byte[] dest = new byte[src.Length];
            using (var aes = new AesCryptoServiceProvider())
            {
                //Parameters of encoding
                aes.BlockSize = 128;
                aes.KeySize = 128;
                aes.IV = Encoding.UTF8.GetBytes(AesIV);
                aes.Key = Encoding.UTF8.GetBytes(AesKey);
                aes.Mode = CipherMode.CBC;
                using (ICryptoTransform encrypt = aes.CreateEncryptor(aes.Key, aes.IV))
                {
                    //Returns encrypted values
                    return encrypt.TransformFinalBlock(src, 0, src.Length);
                }
            }
        }

        //Returns ciphertext
        public string EncryptString(string toEncrypt)
        {
            return Convert.ToBase64String(EncryptBytes(toEncrypt));
        }

        //AES method to decrypt
        public string DecryptToBytesUsingCBC(byte[] toDecrypt)
        {
            //Gets bytes of string to encrypt
            byte[] src = toDecrypt;
            byte[] dest = new byte[src.Length];
            using (var aes = new AesCryptoServiceProvider())
            {
                //Parameters of encoding
                aes.BlockSize = 128;
                aes.KeySize = 128;
                aes.IV = Encoding.UTF8.GetBytes(AesIV);
                aes.Key = Encoding.UTF8.GetBytes(AesKey);
                aes.Mode = CipherMode.CBC;
                using (ICryptoTransform decrypt = aes.CreateDecryptor(aes.Key, aes.IV))
                {
                    //Returns encrypted values
                    byte[] decryptedText = decrypt.TransformFinalBlock(src, 0, src.Length);
                    return Encoding.UTF8.GetString(decryptedText);
                }
            }
        }

        //Returns plaintext
        public string DecryptUsingCBC(string toDecrypt)
        {
            return DecryptToBytesUsingCBC(Convert.FromBase64String(toDecrypt));
        }
    }
}
