//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;

//namespace SolidEdge.Community.AddIn
//{
//    public enum EdgeBarPageDocumentTypes
//    {
//        Assembly,
//        Draft,
//        Part,
//        SheetMetal
//    }

//    [System.AttributeUsage(System.AttributeTargets.Class, AllowMultiple=true)]
//    public class EdgeBarPageAttribute : System.Attribute
//    {
//        private EdgeBarPageDocumentTypes _documentType;
//        private int _imageId;

//        public EdgeBarPageAttribute(EdgeBarPageDocumentTypes documentType, int imageId)
//        {
//            _documentType = documentType;
//            _imageId = imageId;
//        }

//        public EdgeBarPageDocumentTypes DocumentType { get { return _documentType; } }
//        public int ImageId { get { return _imageId; } }
//    }
//}