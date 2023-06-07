using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace WServiceIISM3
{
    internal class RijndaelAlgorithm
    {
        public static string Decrypt
        (
            string cipherText,
            string passPhrase,
            string saltValue,
            string hashAlgorithm,
            int passwordIterations,
            string initVector,
            int keySize
        )
        {
            //Преобразование строк, определяющих характеристики ключа шифрования, в байтовые массивы.
            byte[] initVectorBytes = Encoding.ASCII.GetBytes(initVector);
            byte[] saltValueBytes = Encoding.ASCII.GetBytes(saltValue);

            //Преобразование нашего зашифрованного текста в массив байтов.
            byte[] cipherTextBytes = Convert.FromBase64String(cipherText);

            //Во-первых, мы должны создать пароль, из которого будет получен ключ
            //Этот пароль будет создан из указанной парольной фразы и значения соли.
            //Пароль будет создан с использованием указанного алгоритма хэша. Создание пароля может выполняться в нескольких итерациях.
            PasswordDeriveBytes password = new PasswordDeriveBytes
            (
                passPhrase,
                saltValueBytes,
                hashAlgorithm,
                passwordIterations
            );

            //Используйте пароль для создания псевдослучайных байтов для шифрования
            //ключа. Укажите размер ключа в байтах (вместо битов).
            byte[] keyBytes = password.GetBytes(keySize / 8);

            //Создать неинициализированный объект шифрования Rijndael.
            RijndaelManaged symmetricKey = new RijndaelManaged
            {
                //Целесообразно установить режим шифрования "Цепочка блоков шифрования"
                //(CBC). Используйте параметры по умолчанию для других симметричных ключевых параметров.
                Mode = CipherMode.CBC
            };

            //Создать дешифратор из существующих байтов ключа и инициализации
            //вектора. Размер ключа определяется на основе номера ключа
            //байты.
            ICryptoTransform decryptor = symmetricKey.CreateDecryptor
            (
                keyBytes,
                initVectorBytes
            );

            //Определите поток памяти, который будет использоваться для хранения зашифрованных данных.
            System.IO.MemoryStream memoryStream = new MemoryStream(cipherTextBytes);

            //Определите криптографический поток (всегда используйте режим чтения для шифрования).
            CryptoStream cryptoStream = new CryptoStream
            (
                memoryStream,
                decryptor,
                CryptoStreamMode.Read
            );
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];

            //Начните расшифровывать.
            int decryptedByteCount = cryptoStream.Read
            (
                plainTextBytes,
                0,
                plainTextBytes.Length
            );

            //Закройте оба потока.
            memoryStream.Close();
            cryptoStream.Close();

            //Преобразование расшифрованных данных в строку.
            //Предположим, что исходная строка открытого текста была UTF8-encoded.
            string plainText = Encoding.UTF8.GetString
            (
                plainTextBytes,
                0,
                decryptedByteCount
            );

            //Возвратите расшифрованную последовательность.  
            return plainText;
        }
    }
}
