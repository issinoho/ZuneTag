// ------------------------------------------------------------------
//  DrunkenBakery Zune Tag
//  ZuneTag.ZuneTag
// 
//  <copyright file="Attribute.cs" company="The Drunken Bakery">
//      Copyright (c) 2009-2012 The Drunken Bakery. All rights reserved.
//  </copyright>
// 
//  Author: IRS
// ------------------------------------------------------------------
namespace DrunkenBakery.ZuneTag
{
    using WMFSDKWrapper;

    /// <summary>
    /// The attribute.
    /// </summary>
    internal class Attribute
    {
        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="Attribute"/> class.
        /// </summary>
        /// <param name="index">
        /// The _index.
        /// </param>
        /// <param name="name">
        /// The _name.
        /// </param>
        /// <param name="value">
        /// The _value.
        /// </param>
        /// <param name="type">
        /// The _type.
        /// </param>
        public Attribute(ushort index, string name, string value, WMT_ATTR_DATATYPE type)
        {
            this.Index = index;
            this.Name = name;
            this.Value = value;
            this.Type = type;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets or sets the index.
        /// </summary>
        public ushort Index { get; set; }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the type.
        /// </summary>
        public WMT_ATTR_DATATYPE Type { get; set; }

        /// <summary>
        /// Gets or sets the value.
        /// </summary>
        public string Value { get; set; }

        #endregion

        #region Public Methods and Operators

        /// <summary>
        /// The to string.
        /// </summary>
        /// <returns>
        /// The <see cref="string"/>.
        /// </returns>
        public override string ToString()
        {
            return this.Name + " = " + this.Value + " (" + this.Type + ")";
        }

        #endregion
    }
}