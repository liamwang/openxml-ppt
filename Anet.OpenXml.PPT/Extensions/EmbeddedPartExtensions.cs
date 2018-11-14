using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;

namespace Anet.OpenXml.PPT
{
    public static class EmbeddedPartExtensions
    {
        /// <summary>
        /// 用Base64替换嵌入文件数据
        /// </summary>
        /// <param name="base64">Base64</param>
        public static void ReplaceWithBase64(this EmbeddedPackagePart embeddedPackagePart, string base64)
        {
            if (string.IsNullOrWhiteSpace(base64)) return;

            using (var data = GetBinaryDataStream(base64))
            {
                embeddedPackagePart.FeedData(data);
            }
        }

        private static Stream GetBinaryDataStream(string base64String)
        {
            return new MemoryStream(Convert.FromBase64String(base64String));
        }
    }
}
