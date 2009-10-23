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
    
    class Attribute
    {
        ushort _index;
        string _name;
        string _value;
        WMT_ATTR_DATATYPE _type;

        public Attribute(ushort _index, string _name, string _value, WMT_ATTR_DATATYPE _type)
        {
            this._index = _index;
            this._name = _name;
            this._value = _value;
            this._type = _type;
        }

        public ushort Index
        {
            get { return _index; }
            set { _index = value; }
        }

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }

        public string Value
        {
            get { return _value; }
            set { _value = value; }
        }

        public WMT_ATTR_DATATYPE Type
        {
            get { return _type; }
            set { _type = value; }
        }

        public override string ToString()
        {
            return this.Name + " = " + this.Value + " (" + this.Type + ")";
        }
    }
}
