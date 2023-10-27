//------------------------------------------------------------------
// Zune Meta Tag Editor
// Attribute Class
//
// <copyright file="Attribute.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Holds tag attributes.
//
// Author: IRS
// $Revision: 1.2 $
//------------------------------------------------------------------using System;

namespace DrunkenBakery.ZuneTag
{
    using WMFSDKWrapper;

    internal class Attribute
    {
        public Attribute(ushort index, string name, string value, WMT_ATTR_DATATYPE type)
        {
            this.Index = index;
            this.Name = name;
            this.Value = value;
            this.Type = type;
        }

        public ushort Index
        {
            get;
            set;
        }

        public string Name
        {
            get;
            set;
        }

        public string Value
        {
            get;
            set;
        }

        public WMT_ATTR_DATATYPE Type
        {
            get;
            set;
        }

        public override string ToString()
        {
            return this.Name + " = " + this.Value + " (" + this.Type + ")";
        }
    }
}
