//------------------------------------------------------------------
// Zune Meta Tag Editor
// TMDB Search Result Class
//
// <copyright file="TMDbSearchResult.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Holds a search result entry from TMDB.
//
// Author: IRS
// $Revision: 1.2 $
//------------------------------------------------------------------

namespace DrunkenBakery.ZuneTag
{
    using TMDbLib.Objects.Search;

    internal class TMDbSearchResult
    {
        public TMDbSearchResult()
        {
            this.ExtraData = false;
            this.Genre = "unknown";
            this.Director = "unknown";
        }

        public SearchMovie Movie
        {
            get;
            set;
        }

        public bool ExtraData
        {
            get;
            set;
        }

        public string Genre
        {
            get;
            set;
        }

        public string Url
        {
            get;
            set;
        }

        public string Director
        {
            get;
            set;
        }

        public override string ToString()
        {
            return this.Movie.Title;
        }
    }
}