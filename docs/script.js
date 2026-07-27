// Enhances the static "latest release" download links/badges with the real
// current version, pulled live from the GitHub Releases API. The links
// already work without this (releases/latest/download/... always resolves
// to the newest asset), so a failed or slow fetch just leaves the fallback
// text in place.
(function () {
  fetch("https://api.github.com/repos/issinoho/ZuneTag/releases/latest")
    .then(function (response) {
      if (!response.ok) {
        throw new Error("GitHub API returned " + response.status);
      }

      return response.json();
    })
    .then(function (release) {
      var badges = document.querySelectorAll("#version-badge, #version-badge-2");
      badges.forEach(function (badge) {
        badge.textContent = release.tag_name;
      });

      var asset = (release.assets || []).find(function (a) {
        return a.name === "ZuneTag.exe";
      });

      if (asset) {
        var downloadLink = document.getElementById("primary-download");
        if (downloadLink) {
          downloadLink.href = asset.browser_download_url;
        }
      }
    })
    .catch(function () {
      // Leave the static fallback text/links as-is.
    });
})();
