// ------------------------------------------------------------------
//  DrunkenBakery Zune Tag
//  ZuneTag.ZuneTag
// 
//  <copyright file="AmazonEntry.cs" company="The Drunken Bakery">
//      Copyright (c) 2009-2012 The Drunken Bakery. All rights reserved.
//  </copyright>
// 
//  Author: IRS
// ------------------------------------------------------------------
namespace DrunkenBakery.ZuneTag
{
    /// <summary>
    /// The amazon entry.
    /// </summary>
    internal class AmazonEntry
    {
        #region Constructors and Destructors

        /// <summary>
        /// Initializes a new instance of the <see cref="AmazonEntry"/> class.
        /// </summary>
        /// <param name="asin">
        /// The _asin.
        /// </param>
        /// <param name="title">
        /// The _title.
        /// </param>
        /// <param name="director">
        /// The _director.
        /// </param>
        /// <param name="date">
        /// The _date.
        /// </param>
        /// <param name="description">
        /// The _description.
        /// </param>
        /// <param name="url">
        /// The _ url.
        /// </param>
        /// <param name="genre">
        /// The _genre.
        /// </param>
        /// <param name="rating">
        /// The _rating.
        /// </param>
        /// <param name="cover">
        /// The _cover.
        /// </param>
        public AmazonEntry(
            string asin, 
            string title, 
            string director, 
            string date, 
            string description, 
            string url, 
            string genre, 
            string rating, 
            string cover)
        {
            this.Asin = asin;
            this.Title = title;
            this.Director = director;
            this.Date = date;
            this.Year = date.Length >= 4 ? date.Substring(0, 4) : string.Empty;
            this.Description = description;
            this.Url = url;
            this.Genre = genre;
            this.Rating = rating;
            this.Cover = cover;
        }

        #endregion

        #region Public Properties

        /// <summary>
        /// Gets or sets the asin.
        /// </summary>
        public string Asin { get; set; }

        /// <summary>
        /// Gets or sets the cover.
        /// </summary>
        public string Cover { get; set; }

        /// <summary>
        /// Gets or sets the date.
        /// </summary>
        public string Date { get; set; }

        /// <summary>
        /// Gets or sets the description.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the director.
        /// </summary>
        public string Director { get; set; }

        /// <summary>
        /// Gets or sets the genre.
        /// </summary>
        public string Genre { get; set; }

        /// <summary>
        /// Gets or sets the rating.
        /// </summary>
        public string Rating { get; set; }

        /// <summary>
        /// Gets or sets the title.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Gets or sets the url.
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// Gets or sets the year.
        /// </summary>
        public string Year { get; set; }

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
            return this.Title;
        }

        #endregion
    }
}