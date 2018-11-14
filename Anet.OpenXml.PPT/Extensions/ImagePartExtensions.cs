using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

namespace Anet.OpenXml.PPT
{
    public static class ImagePartExtensions
    {
        /// <summary>
        /// 用Base64替换图片数据
        /// </summary>
        /// <param name="base64">Base64</param>
        public static void ReplaceWithBase64(this ImagePart imagePart, string base64)
        {
            if (string.IsNullOrWhiteSpace(base64)) return;

            using (var data = GetBinaryDataStream(base64))
            {
                imagePart.FeedData(data);
            }
        }

        private static Stream GetBinaryDataStream(string base64String)
        {
            return new MemoryStream(Convert.FromBase64String(base64String));
        }
    }
}
