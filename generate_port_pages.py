import os
import re
import openpyxl

# Configuration Paths
REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
PORTS_DIR = os.path.join(REPO_ROOT, "ports")
CSS_FILE = os.path.join(REPO_ROOT, "assets", "css", "port-pages.css")
EXCEL_PATH = r"D:\FalconXoft Ship Chandler\Website Data\Ship_Chandler_SEO_49_Pages_FULL.xlsx"

PORT_PAGE_TEMPLATE = """<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>__META_TITLE__</title>
  <meta name="description"
    content="__META_DESC__">
  <meta name="keywords"
    content="ship chandler, ship supply, provisions, marine stores, bonded stores, deck stores, engine stores, vessel supply, marine procurement, __PAGE_NAME__">
  <meta name="author" content="Ship Chandler Germany CO., LTD">
  <link rel="canonical" href="https://shipchandlergermany.com/ports/__SLUG__">

  <!-- Open Graph Meta Tags -->
  <meta property="og:title" content="__META_TITLE__">
  <meta property="og:description"
    content="__META_DESC__">
  <meta property="og:type" content="website">
  <meta property="og:url" content="https://shipchandlergermany.com/ports/__SLUG__">
  <meta property="og:site_name" content="Ship Chandler Germany">

  <!-- Twitter Card Meta Tags -->
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="__META_TITLE__">
  <meta name="twitter:description"
    content="__META_DESC__">

  <!-- Google Fonts: Editorial + Monospace -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link
    href="https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=JetBrains+Mono:wght@300;400;500;600&family=Manrope:wght@300;400;500;600;700;800&display=swap"
    rel="stylesheet">

  <!-- Font Awesome -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css">

  <!-- Favicon -->
  <link rel="icon" type="image/jpeg" href="../assets/imgs/favicon.jpg">
  <link rel="shortcut icon" href="../favicon.jpg">
  <link rel="apple-touch-icon" href="../assets/imgs/favicon.jpg">

  <!-- Centralized Port Pages CSS -->
  <link rel="stylesheet" href="../assets/css/port-pages.css">
</head>

<body>

  <!-- TOP ANNOUNCEMENT BAR -->
  <div class="top-bar">
    <div class="container">
      <span>Port of Hamburg · Germany · 24/7 Vessel Supply</span>
      <span>Direct WhatsApp: <a href="https://wa.me/6590190005" target="_blank"
          rel="noopener noreferrer">+65-9019-0005</a></span>
    </div>
  </div>

  <!-- HEADER / NAVBAR -->
  <header class="site-header" id="site-header">
    <div class="container header-inner">
      <a class="logo" href="../index.html">
        <img src="../assets/imgs/logo-shipchandler-germany.png" alt="ShipChandler Germany Logo" width="44" height="44" onerror="this.onerror=null;this.src='../logo-shipchandler-germany.png';">
        <span class="logo-text">
          <strong>ShipChandler</strong>
          <span>Germany</span>
        </span>
      </a>

      <button class="nav-toggle" type="button" aria-label="Toggle navigation">
        <span></span><span></span><span></span>
      </button>

      <nav class="primary-nav" id="primary-nav">
        <ul class="nav-list">
          <li class="nav-item"><a class="nav-link" href="../index.html">Home</a></li>
          <li class="nav-item"><a class="nav-link" href="../about.html">About</a></li>

          <!-- Services Dropdown -->
          <li class="nav-item has-dropdown">
            <a class="nav-link" href="../services.html">Services <span class="nav-arrow"><svg viewBox="0 0 24 24"
                  width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"
                  stroke-linejoin="round">
                  <polyline points="6 9 12 15 18 9"></polyline>
                </svg></span></a>
            <div class="dropdown-menu">
              <a href="../services.html#provisions">Provisions Supply</a>
              <a href="../services.html#bonded">Bonded Stores</a>
              <a href="../services.html#deck">Deck Stores</a>
              <a href="../services.html#engine">Engine Stores</a>
              <a href="../services.html#cabin">Cabin & Galley</a>
              <a href="../services.html#safety">Safety & PPE</a>
              <a href="../services.html#electrical">Electrical Stores</a>
              <a href="../services.html#chemicals">Marine Chemicals</a>
              <a href="../services.html#welding">Welding Equipment</a>
              <a href="../services.html#spare-parts">Spare Parts</a>
              <a href="../services.html#publications">Nautical Charts</a>
              <a href="../services.html#tools">Tools & Hardware</a>
              <a href="../services.html#urgent">Urgent Supply</a>
            </div>
          </li>

          <!-- Ports Mega Menu -->
          <li class="nav-item has-dropdown">
            <a class="nav-link" href="javascript:void(0)">Ports <span class="nav-arrow"><svg viewBox="0 0 24 24"
                  width="12" height="12" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"
                  stroke-linejoin="round">
                  <polyline points="6 9 12 15 18 9"></polyline>
                </svg></span></a>
            <div class="mega-menu">
              <div class="mega-header">
                <div class="mega-header-title">Global Port Network</div>
                <div class="mega-search">
                  <i class="fas fa-search" style="color:var(--ink-soft);"></i>
                  <input type="text" id="mega-port-search" placeholder="Search 49 ports...">
                </div>
              </div>
              <div class="mega-grid" id="mega-port-list">
                <!-- Populated via JS -->
              </div>
            </div>
          </li>

          <li class="nav-item"><a class="nav-link" href="../contact.html">Contact</a></li>
        </ul>
        <a class="btn btn-primary" style="padding: 10px 24px; font-size: 0.68rem;" href="../request-quote.html">Get
          Quote</a>
      </nav>
    </div>
  </header>

  <main>
    <!-- SUBPAGE HERO -->
    <section class="subpage-hero">
      <div class="subpage-hero-split">
        <div class="subpage-hero-left">
          <div class="container-inner">
            <div class="breadcrumb">
              <a href="../index.html">Home</a>
              <span>/</span>
              <a href="../index.html#ports">Ports</a>
              <span>/</span>
              <span>__PAGE_NAME__</span>
            </div>
            <h1 class="subpage-hero-title">
              __HERO_TITLE__
            </h1>
            <p class="subpage-hero-desc">__META_DESC__</p>
          </div>
        </div>
        <div class="subpage-hero-right">
          <img src="../assets/imgs/germany-hamburg-ships.jpg"
            alt="__H1__" onerror="this.onerror=null;this.src='../germany-hamburg-ships.jpg';">
          <div class="hero-side-label">__HERO_SIDE_LABEL__</div>
        </div>
      </div>
    </section>

    <!-- PORT CONTENT SECTION -->
    <section class="port-content">
      <div class="container">
        <article class="prose">
__ARTICLE_HTML__
        </article>
      </div>
    </section>

    <!-- CTA SECTION -->
    <section class="cta-section">
      <div class="container">
        <h2>Ready to Supply Your Vessel?</h2>
        <p>Send us your requisition list and our team will provide a competitive quote within 2 hours.</p>
        <a class="btn btn-primary" href="../request-quote.html">Request a Quote Now</a>
      </div>
    </section>
  </main>

  <!-- FOOTER -->
  <footer class="site-footer">
    <div class="container">
      <div class="footer-top">
        <div class="footer-brand">
          <a class="logo" href="../index.html">
            <img src="../assets/imgs/logo-shipchandler-germany.png" alt="Logo" width="44" height="44" onerror="this.onerror=null;this.src='../logo-shipchandler-germany.png';">
            <span class="logo-text">
              <strong>ShipChandler</strong>
              <span style="color: var(--acid);">Germany</span>
            </span>
          </a>
          <p>SHIP CHANDLER GERMANY CO., LTD is a licensed marine supplier providing provisions, technical stores, bonded
            goods and emergency vessel supply across all German ports and anchorages.</p>
        </div>

        <div class="footer-col">
          <h4>Quick Links</h4>
          <ul class="footer-links">
            <li><a href="../index.html">Home</a></li>
            <li><a href="../about.html">About Us</a></li>
            <li><a href="../services.html">Services</a></li>
            <li><a href="../contact.html">Contact</a></li>
            <li><a href="../request-quote.html">Request a Quote</a></li>
          </ul>
        </div>

        <div class="footer-col">
          <h4>Services</h4>
          <ul class="footer-links">
            <li><a href="../services.html#provisions">Provisions Supply</a></li>
            <li><a href="../services.html#bonded">Bonded Stores</a></li>
            <li><a href="../services.html#deck">Deck Stores</a></li>
            <li><a href="../services.html#engine">Engine Stores</a></li>
            <li><a href="../services.html#safety">Safety & PPE</a></li>
          </ul>
        </div>

        <div class="footer-col">
          <h4>Contact Us</h4>
          <ul class="footer-contact">
            <li><i class="fas fa-map-marker-alt"></i> Port of Hamburg, Am Strandkai 1, 20457 Hamburg, Germany</li>
            <li><i class="fas fa-envelope"></i> <a
                href="mailto:sales@shipchandlergermany.com">sales@shipchandlergermany.com</a></li>
            <li><i class="fas fa-user-tie"></i> Thomas Müller</li>
            <li>
              <a class="hero-whatsapp-bar" href="https://wa.me/6590190005" target="_blank" rel="noopener noreferrer"
                aria-label="Direct WhatsApp contact (+65-9019-0005)">
                <span class="hero-whatsapp-icon" aria-hidden="true">
                  <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
                    <path
                      d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z" />
                  </svg>
                </span>
                <span class="hero-whatsapp-tag">WhatsApp:</span>
                <span class="hero-whatsapp-num">(+65-9019-0005)</span>
                <span class="hero-whatsapp-arrow" aria-hidden="true">&rarr;</span>
              </a>
            </li>
          </ul>
        </div>
      </div>

      <div class="footer-bottom">
        <div>&copy; 2026 SHIP CHANDLER GERMANY CO., LTD. All rights reserved.</div>
      </div>
    </div>
  </footer>

  <!-- FLOATING WHATSAPP -->
  <a href="https://wa.me/6590190005" class="whatsapp-float" target="_blank" rel="noopener noreferrer"
    aria-label="WhatsApp Contact +65-9019-0005">
    <i class="fab fa-whatsapp"></i>
  </a>

  <script>
    // --- Port List Data for Dropdown ---
    const ports = [
      "Ship Chandler Dar Es Salaam", "Ship Chandler Apapa", "Ship Chandler Onne Port Nigeria", "San Pedro Ship Supply Chandler",
      "Ship Chandler Abidjan", "Ship Chandler Freetown", "Monrovia Ship Chandler", "Ship Chandler Maputo", "Ship Chandler Matadi",
      "Ship Chandler Beira", "Ship Chandler Angola", "Ship Chandler Lobito", "ShipChandler Lome", "Shipchandler Lome",
      "Ivory Coast Ship Chandler", "Nouadhibou Ship Chandler", "Ship Chandler Luanda", "Ship Chandler Cape Verde",
      "Ship Chandler Gabon", "Ship Chandler Namibia", "Walvis Bay Ship Chandler", "Ship Chandler Gentil Port - Gabon",
      "Ship Chandler Takoradi", "Ship Chandler Congo", "Shipchandler Lagos", "Ship Chandler Port Ambon",
      "Shipchandler Pointe Noire", "Ship Chandler Djibouti", "Ship Chandler Cabinda", "Ship Chandler Tema",
      "Ship Chandler Mindelo", "Ship Chandler Casablanca", "Ship Chandler Safi Port Morocco", "Ship Chandler Kamsar",
      "Ship Chandler Soyo", "Dakar Ship Chandler", "Ship Chandler Senegal", "Ship Chandler Dakar", "Bissau Ship Chandler - Guinea",
      "Conakry Ship Chandler", "Ship Chandler Cotonou", "Douala Ship Chandler", "Liberia Ship Chandler", "Namibia Ship Chandler",
      "Ship Chandler Dakar", "Ship Chandler Mombasa", "Ship Chandlers Saldanha Bay", "Ship Chandler Owendo", "Ship Chandlers Richards Bay"
    ];

    const megaList = document.getElementById('mega-port-list');
    if (megaList) {
      const third = Math.ceil(ports.length / 3);
      const col1 = ports.slice(0, third);
      const col2 = ports.slice(third, third * 2);
      const col3 = ports.slice(third * 2);

      function createPortGroup(title, items) {
        const group = document.createElement('div');
        group.className = 'mega-group';
        group.innerHTML = `<h4>${title}</h4>`;
        items.forEach(port => {
          const a = document.createElement('a');
          a.href = `${port.toLowerCase().replace(/[^a-z0-9]+/g, '-')}.html`;
          a.textContent = port;
          group.appendChild(a);
        });
        return group;
      }

      megaList.appendChild(createPortGroup('African Ports (A-D)', col1));
      megaList.appendChild(createPortGroup('African Ports (E-M)', col2));
      megaList.appendChild(createPortGroup('African Ports (N-Z)', col3));

      // Mega Menu Search
      const searchInput = document.getElementById('mega-port-search');
      if (searchInput) {
        searchInput.addEventListener('input', function (e) {
          const query = e.target.value.toLowerCase();
          const allLinks = megaList.querySelectorAll('a');
          const groups = megaList.querySelectorAll('.mega-group');
          allLinks.forEach(link => {
            const text = link.textContent.toLowerCase();
            link.style.display = text.includes(query) ? 'block' : 'none';
          });
          groups.forEach(group => {
            const hasVisible = Array.from(group.querySelectorAll('a')).some(a => a.style.display !== 'none');
            group.style.display = hasVisible ? 'block' : 'none';
          });
        });
      }
    }

    const megaMenuEl = document.querySelector('.mega-menu');
    const siteHeaderEl = document.getElementById('site-header');

    function positionMegaMenu() {
      if (!megaMenuEl || !siteHeaderEl) return;
      const headerRect = siteHeaderEl.getBoundingClientRect();
      // Anchor the mega menu flush against the bottom of the sticky header,
      // with a 4px gap so it doesn't visually touch.
      megaMenuEl.style.top = (headerRect.bottom + 4) + 'px';
    }

    // Run on load, on scroll (in case the top-bar shrinks on scroll), and on resize
    window.addEventListener('load', positionMegaMenu);
    window.addEventListener('scroll', positionMegaMenu, { passive: true });
    window.addEventListener('resize', positionMegaMenu);

    // Sticky Header
    const header = document.querySelector('.site-header');
    window.addEventListener('scroll', () => {
      header.classList.toggle('scrolled', window.scrollY > 20);
    });

    // Mobile Nav Toggle
    const navToggle = document.querySelector('.nav-toggle');
    const primaryNav = document.querySelector('.primary-nav');

    navToggle.addEventListener('click', function () {
      primaryNav.classList.toggle('open');
      const spans = this.querySelectorAll('span');
      if (primaryNav.classList.contains('open')) {
        spans[0].style.transform = 'rotate(45deg) translate(5px, 5px)';
        spans[1].style.opacity = '0';
        spans[2].style.transform = 'rotate(-45deg) translate(5px, -5px)';
      } else {
        spans[0].style.transform = 'none';
        spans[1].style.opacity = '1';
        spans[2].style.transform = 'none';
      }
    });

    // Mobile Dropdown Toggles
    document.querySelectorAll('.nav-item.has-dropdown').forEach(item => {
      const link = item.querySelector('.nav-link');
      link.addEventListener('click', function (e) {
        if (window.innerWidth <= 1024) {
          e.preventDefault();
          item.classList.toggle('open');
        }
      });
    });
  </script>
</body>

</html>"""


def format_hero(page_name):
    """
    Formats the subpage-hero-title and hero-side-label matching Bissau's editorial style.
    """
    raw = str(page_name).strip()
    
    # Determine side label
    clean = re.sub(r'(?i)\bship\s*chandlers?\b', '', raw)
    clean = re.sub(r'(?i)\bship\s*supply\s*chandler\b', '', clean)
    clean = clean.replace('-', '·')
    parts = [p.strip() for p in re.split(r'[\s·]+', clean) if p.strip()]
    
    if 'onne' in raw.lower() and 'nigeria' in raw.lower():
        side_label = 'Port Page · Onne Port · Nigeria'
    elif 'safi' in raw.lower() and 'morocco' in raw.lower():
        side_label = 'Port Page · Safi Port · Morocco'
    elif 'gentil' in raw.lower() and 'gabon' in raw.lower():
        side_label = 'Port Page · Port Gentil · Gabon'
    elif 'bissau' in raw.lower() and 'guinea' in raw.lower():
        side_label = 'Port Page · Bissau · Guinea'
    elif 'dar' in raw.lower() and 'salaam' in raw.lower():
        side_label = 'Port Page · Dar Es Salaam'
    elif 'cape' in raw.lower() and 'verde' in raw.lower():
        side_label = 'Port Page · Cape Verde'
    elif 'walvis' in raw.lower() and 'bay' in raw.lower():
        side_label = 'Port Page · Walvis Bay'
    elif 'richards' in raw.lower() and 'bay' in raw.lower():
        side_label = 'Port Page · Richards Bay'
    elif 'saldanha' in raw.lower() and 'bay' in raw.lower():
        side_label = 'Port Page · Saldanha Bay'
    elif 'ivory' in raw.lower() and 'coast' in raw.lower():
        side_label = 'Port Page · Ivory Coast'
    elif 'pointe' in raw.lower() and 'noire' in raw.lower():
        side_label = 'Port Page · Pointe Noire'
    elif 'san' in raw.lower() and 'pedro' in raw.lower():
        side_label = 'Port Page · San Pedro'
    elif 'port' in raw.lower() and 'ambon' in raw.lower():
        side_label = 'Port Page · Port Ambon'
    else:
        side_label = 'Port Page · ' + ' '.join(parts)

    # Determine subpage-hero-title with editorial typography
    if 'bissau' in raw.lower() and 'guinea' in raw.lower():
        hero_title = 'Bissau<br><em>Ship Chandler</em><br>Guinea'
    elif ' - ' in raw:
        p1, p2 = [p.strip() for p in raw.split(' - ', 1)]
        p1_clean = re.sub(r'(?i)\bship\s*chandlers?\b', '', p1).strip()
        hero_title = f'{p1_clean}<br><em>Ship Chandler</em><br>{p2}'
    elif re.search(r'(?i)^ship\s*chandlers?\s+', raw):
        loc = re.sub(r'(?i)^ship\s*chandlers?\s+', '', raw).strip()
        hero_title = f'<em>Ship Chandler</em><br>{loc}'
    elif re.search(r'(?i)^shipchandler\s+', raw):
        loc = re.sub(r'(?i)^shipchandler\s+', '', raw).strip()
        hero_title = f'<em>Ship Chandler</em><br>{loc}'
    elif re.search(r'(?i)\s+ship\s*supply\s*chandler$', raw):
        loc = re.sub(r'(?i)\s+ship\s*supply\s*chandler$', '', raw).strip()
        hero_title = f'{loc}<br><em>Ship Supply Chandler</em>'
    elif re.search(r'(?i)\s+ship\s*chandler$', raw):
        loc = re.sub(r'(?i)\s+ship\s*chandler$', '', raw).strip()
        hero_title = f'{loc}<br><em>Ship Chandler</em>'
    else:
        hero_title = f'<em>Ship Chandler</em><br>{raw}'

    return hero_title, side_label


def render_article_html(desc_text):
    """
    Parses the Markdown description from Excel into semantic HTML
    matching the exact structure and indentation of bissau-ship-chandler-guinea.html
    """
    blocks = [b.strip() for b in desc_text.split('\n\n') if b.strip()]
    lines = []
    
    for b in blocks:
        if b.startswith('# '):
            # Skip top-level H1 (rendered in Hero section)
            continue
        elif b.startswith('## '):
            heading_text = b[3:].strip()
            lines.append(f'          <h2>{heading_text}</h2>')
        elif b.startswith('- '):
            lines.append('          <ul>')
            for item in b.split('\n'):
                item = item.strip()
                if not item:
                    continue
                val = re.sub(r'^\s*-\s*', '', item)
                lines.append(f'            <li>{val}</li>')
            lines.append('          </ul>')
        else:
            # Inline bold markdown replacement
            content = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', b)
            lines.append(f'          <p>{content}</p>')
                
    return '\n'.join(lines)


def build_port_page_html(page_name, h1, meta_title, meta_desc, full_desc, slug):
    """
    Builds the full HTML document for a port page matching Bissau line-for-line
    using centralized port-pages.css and standardized JPG assets with fallbacks.
    """
    hero_title_html, hero_side_label = format_hero(page_name)
    article_html = render_article_html(full_desc)

    html = PORT_PAGE_TEMPLATE
    html = html.replace('__META_TITLE__', str(meta_title or ''))
    html = html.replace('__META_DESC__', str(meta_desc or ''))
    html = html.replace('__PAGE_NAME__', str(page_name or ''))
    html = html.replace('__H1__', str(h1 or ''))
    html = html.replace('__SLUG__', slug)
    html = html.replace('__HERO_TITLE__', hero_title_html)
    html = html.replace('__HERO_SIDE_LABEL__', hero_side_label)
    html = html.replace('__ARTICLE_HTML__', article_html)

    return html


def main():
    print(f"Reading Excel data from: {EXCEL_PATH}")
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    sheet = wb['SEO Pages']
    
    rows = list(sheet.iter_rows(values_only=True))[1:]  # Skip header row
    print(f"Found {len(rows)} port rows in Excel.")
    
    os.makedirs(PORTS_DIR, exist_ok=True)
    
    generated_count = 0
    unique_slugs = set()
    
    for row in rows:
        num, page_name, h1, meta_title, meta_desc, full_desc = row[:6]
        if not page_name:
            continue
            
        slug = re.sub(r'[^a-z0-9]+', '-', str(page_name).lower()).strip('-') + '.html'
        out_path = os.path.join(PORTS_DIR, slug)
        
        html_content = build_port_page_html(
            page_name=page_name,
            h1=h1,
            meta_title=meta_title,
            meta_desc=meta_desc,
            full_desc=full_desc,
            slug=slug
        )
        
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
            
        generated_count += 1
        unique_slugs.add(slug)
        print(f"[{num:02d}/49] Generated: {slug} ({page_name})")
        
    print("\n" + "="*60)
    print(f"SUCCESS: Processed {generated_count} rows across {len(unique_slugs)} unique port page files.")
    print("Centralized CSS: assets/css/port-pages.css")
    print("Standardized Images: germany-hamburg-ships.jpg, logo-shipchandler-germany.png, favicon.jpg")
    print("="*60)


if __name__ == "__main__":
    main()
