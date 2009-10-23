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

namespace DrunkenBakery.ZuneTag
{
    class AmazonEntry
    {
        string _asin;
        string _title;
        string _director;
        string _year;
        string _date;
        string _genre;
        string _rating;
        string _description;
        string _URL;
        string _cover;

        public AmazonEntry(string _asin, string _title, string _director, string _date, string _description, string _URL, string _genre, string _rating, string _cover)
        {
            this._asin = _asin;
            this._title = _title;
            this._director = _director;
            this._date = _date;
            this._year = _date.Length >= 4 ? _date.Substring(0,4) : "";
            this._description = _description;
            this._URL = _URL;
            this._genre = _genre;
            this._rating = _rating;
            this._cover = _cover;
        }

        public string Cover
        {
            get { return _cover; }
            set { _cover = value; }
        }

        public string ASIN
        {
            get { return _asin; }
            set { _asin = value; }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; }
        }

        public string Director
        {
            get { return _director; }
            set { _director = value; }
        }

        public string Year
        {
            get { return _year; }
            set { _year = value; }
        }

        public string Date
        {
            get { return _date; }
            set { _date = value; }
        }

        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }

        public string URL
        {
            get { return _URL; }
            set { _URL = value; }
        }

        public string Genre
        {
            get { return _genre; }
            set { _genre = value; }
        }

        public string Rating
        {
            get { return _rating; }
            set { _rating = value; }
        }

        public override string ToString()
        {
            return this.Title;
        }
    }
}
