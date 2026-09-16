=========================================================
  thetaleofher — READY TO DEPLOY ON NETLIFY
=========================================================

FASTEST WAY (drag & drop, ~1 minute)
  1. Go to  https://app.netlify.com/drop
  2. Drag THIS WHOLE FOLDER (or the .zip) onto the page.
  3. Done — Netlify gives you a live link like
     https://random-name-123.netlify.app
  You will NOT see the "No index.html file" warning anymore,
  because index.html is now the homepage.

  Tip: to rename the site, open
  Site configuration > Site details > Change site name.

WHAT'S IN HERE
  index.html          the homepage   (this is what Netlify was asking for)
  katalog.html        the catalog
  produk.html         the product page
  404.html            branded "page not found" page
  netlify.toml        Netlify settings (clean URLs + security headers)
  _redirects          /catalog and /product shortcuts
  favicon.svg         the little icon in the browser tab
  social-preview.jpg  the picture shown when the link is shared (WhatsApp/IG)
  robots.txt          lets Google index the site
  sitemap.xml         list of pages for Google

AFTER YOUR FIRST DEPLOY (2 small edits)
  Replace  YOUR-SITE.netlify.app  with your real address in:
     - robots.txt
     - sitemap.xml
  Then drag the folder in again to update.

USING YOUR OWN DOMAIN (e.g. thetaleofher.com)
  Netlify > Domain management > Add a domain, then follow the steps.
  HTTPS is added automatically and it's free.

UPDATING THE SITE LATER
  Edit the Excel files in the content kit, then run:
      python build_site.py     (rebuilds the pages)
      python make_deploy.py    (rebuilds this deploy folder)
  Drag the folder to Netlify again — that's it.

NOTE
  The pages load fonts and styling from the internet (CDN),
  which is normal for a live website. All product photos are
  embedded inside the HTML, so there is no images folder to upload.
