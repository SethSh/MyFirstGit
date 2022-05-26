using System;
using System.Drawing;
using SubmissionCollector.Enums;
using SubmissionCollector.Properties;

namespace SubmissionCollector.Extensions
{
    public static class ImageExtensions
    {
        public static Bitmap ToResourceBitmap(this MessageType messageType)
        {
            switch (messageType)
            {
                case MessageType.Documentation: return Resources.Document;
                case MessageType.Log: return Resources.Parchment;
                case MessageType.Information: return Resources.Information;
                case MessageType.Warning: return Resources.Warning;
                case MessageType.Stop: return Resources.Stop;
                case MessageType.Success: return Resources.Success;
                default: throw new ArgumentOutOfRangeException();
            }
        }
    }
}
