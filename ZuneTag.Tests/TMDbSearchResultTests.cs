//------------------------------------------------------------------
// Zune Meta Tag Editor
// TMDB Search Result Tests
//
// <copyright file="TMDbSearchResultTests.cs" company="The Drunken Bakery">
//     Copyright (c) 2009 The Drunken Bakery. All rights reserved.
// </copyright>
//
// Editor to update WMV meta tags for the Zune
// Tests for the TMDbSearchResult class.
//
// Author: IRS
// $Revision: 1.2 $
//------------------------------------------------------------------

namespace DrunkenBakery.ZuneTag.Tests
{
    using DrunkenBakery.ZuneTag;

    using TMDbLib.Objects.Search;

    using Xunit;

    public class TMDbSearchResultTests
    {
        [Fact]
        public void Constructor_SetsDefaults()
        {
            var entry = new TMDbSearchResult();

            Assert.False(entry.ExtraData);
            Assert.Equal("unknown", entry.Genre);
            Assert.Equal("unknown", entry.Director);
            Assert.Null(entry.Url);
        }

        [Fact]
        public void ToString_ReturnsMovieTitle()
        {
            var entry = new TMDbSearchResult { Movie = new SearchMovie { Title = "The Matrix" } };

            Assert.Equal("The Matrix", entry.ToString());
        }
    }
}
