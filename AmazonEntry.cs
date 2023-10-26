//------------------------------------------------------------------
// Zune Meta Tag Editor
// Amazon Entry Class
//
// <copyright file="AmazonEntry.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Holds search entry from Amazon.
//
// Author: IRS
// $Revision: 1.2 $
//------------------------------------------------------------------using System;

using TMDbLib.Objects.Search;

namespace DrunkenBakery.ZuneTag
{
    class AmazonEntry
    {
        public SearchMovie movie { get; set; }

        public bool extraData { get; set; }

        public string Genre { get; set; }

        public string URL { get; set; }

        public string Director { get; set; }

        public AmazonEntry()
        {
            this.extraData = false;
            this.Genre = "unknown";
            this.Director = "unknown";
        }

        public override string ToString()
        {
            return movie.Title;
        }
    }
}
