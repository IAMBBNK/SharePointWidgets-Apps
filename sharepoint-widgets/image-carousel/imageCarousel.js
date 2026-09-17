(function () {
  var CHEVRON_LEFT =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><polyline points="15 6 9 12 15 18"></polyline></svg>';
  var CHEVRON_RIGHT =
    '<svg viewBox="0 0 24 24" aria-hidden="true"><polyline points="9 6 15 12 9 18"></polyline></svg>';

  function onReady(fn) {
    if (document.readyState === "loading") {
      document.addEventListener("DOMContentLoaded", fn);
    } else {
      fn();
    }
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function escapeAttr(value) {
    return escapeHtml(value).replace(/'/g, "&#39;");
  }

  function normalizeItems(items) {
    if (!Array.isArray(items)) {
      return [];
    }
    return items
      .map(function (item) {
        if (!item || !item.image) {
          return null;
        }
        return {
          image: String(item.image),
          title: item.title ? String(item.title) : "",
          description: item.description ? String(item.description) : ""
        };
      })
      .filter(Boolean);
  }

  function preferredVisibleCount(config) {
    return Number(config.visibleCount) === 1 ? 1 : 3;
  }

  function ImageCarousel(root, config) {
    this.root = root;
    this.config = config || {};
    this.items = normalizeItems(this.config.items);
    this.index = 0;
    this.direction = "next";
    this.isOpen = false;
    this.lastFocused = null;
    this.handleKeydown = this.handleKeydown.bind(this);
    this.handleResize = this.handleResize.bind(this);
    this.render();
  }

  ImageCarousel.prototype.getVisibleCount = function () {
    var preferred = preferredVisibleCount(this.config);
    if (window.innerWidth < 720) {
      preferred = 1;
    }
    return Math.max(1, Math.min(preferred, this.items.length || 1));
  };

  ImageCarousel.prototype.render = function () {
    var config = this.config;
    var buttonText = config.buttonText || "View gallery";
    var buttonIcon = config.buttonIcon ? String(config.buttonIcon).trim() : "";
    var buttonColor = config.buttonColor || "#0f69af";
    var align = String(config.buttonAlign || "left").toLowerCase();

    this.root.className = "imgc-widget";
    if (align === "center") {
      this.root.classList.add("is-center");
    } else if (align === "right") {
      this.root.classList.add("is-right");
    }

    if (!this.items.length) {
      this.root.innerHTML = '<div class="imgc-empty">No carousel images configured.</div>';
      return;
    }

    var iconHtml = buttonIcon
      ? '<img class="imgc-trigger-icon" src="' +
        escapeAttr(buttonIcon) +
        '" alt="" />'
      : "";

    this.root.innerHTML =
      '<button type="button" class="imgc-trigger" style="background:' +
      escapeAttr(buttonColor) +
      '">' +
      iconHtml +
      "<span>" +
      escapeHtml(buttonText) +
      "</span></button>";

    this.trigger = this.root.querySelector(".imgc-trigger");
    this.trigger.addEventListener("click", this.open.bind(this));
    this.buildOverlay();
  };

  ImageCarousel.prototype.buildOverlay = function () {
    if (this.overlay) {
      return;
    }

    this.overlay = document.createElement("div");
    this.overlay.className = "imgc-overlay";
    this.overlay.setAttribute("hidden", "hidden");
    this.overlay.innerHTML =
      '<button type="button" class="imgc-close" aria-label="Close gallery">&times;</button>' +
      '<div class="imgc-overlay-inner">' +
      '<div class="imgc-stage">' +
      '<button type="button" class="imgc-nav imgc-prev" aria-label="Previous">' +
      CHEVRON_LEFT +
      "</button>" +
      '<div class="imgc-viewport"><div class="imgc-cards"></div></div>' +
      '<button type="button" class="imgc-nav imgc-next" aria-label="Next">' +
      CHEVRON_RIGHT +
      "</button>" +
      "</div>" +
      '<div class="imgc-dots"></div>' +
      "</div>";

    document.body.appendChild(this.overlay);

    this.cardsEl = this.overlay.querySelector(".imgc-cards");
    this.dotsEl = this.overlay.querySelector(".imgc-dots");
    this.prevBtn = this.overlay.querySelector(".imgc-prev");
    this.nextBtn = this.overlay.querySelector(".imgc-next");
    this.closeBtn = this.overlay.querySelector(".imgc-close");

    this.overlay.addEventListener("click", function (event) {
      if (event.target === this.overlay) {
        this.close();
      }
    }.bind(this));
    this.closeBtn.addEventListener("click", this.close.bind(this));
    this.prevBtn.addEventListener("click", this.prev.bind(this));
    this.nextBtn.addEventListener("click", this.next.bind(this));
  };

  ImageCarousel.prototype.getVisibleItems = function () {
    var count = this.getVisibleCount();
    var visible = [];
    var i;
    for (i = 0; i < count; i += 1) {
      visible.push(this.items[(this.index + i) % this.items.length]);
    }
    return visible;
  };

  ImageCarousel.prototype.renderSlides = function (animate) {
    var visible = this.getVisibleItems();
    var showNav = this.items.length > this.getVisibleCount();

    this.overlay.classList.toggle("is-single", this.getVisibleCount() === 1);
    this.prevBtn.classList.toggle("is-hidden", !showNav);
    this.nextBtn.classList.toggle("is-hidden", !showNav);
    this.dotsEl.classList.toggle("is-hidden", !showNav);

    this.cardsEl.innerHTML = visible
      .map(function (item) {
        return (
          '<article class="imgc-card">' +
          '<div class="imgc-card-image-wrap">' +
          '<img class="imgc-card-image" src="' +
          escapeAttr(item.image) +
          '" alt="' +
          escapeAttr(item.title) +
          '" />' +
          "</div>" +
          '<h3 class="imgc-card-title">' +
          escapeHtml(item.title) +
          "</h3>" +
          '<p class="imgc-card-desc">' +
          escapeHtml(item.description) +
          "</p>" +
          "</article>"
        );
      })
      .join("");

    if (animate) {
      var animClass =
        this.direction === "prev" ? "is-animating-prev" : "is-animating";
      this.cardsEl.classList.remove("is-animating", "is-animating-prev");
      void this.cardsEl.offsetWidth;
      this.cardsEl.classList.add(animClass);
    }

    this.dotsEl.innerHTML = this.items
      .map(function (_item, i) {
        var active = i === this.index ? " is-active" : "";
        return (
          '<button type="button" class="imgc-dot' +
          active +
          '" aria-label="Go to slide ' +
          (i + 1) +
          '"></button>'
        );
      }, this)
      .join("");

    var dots = this.dotsEl.querySelectorAll(".imgc-dot");
    var self = this;
    dots.forEach(function (dot, i) {
      dot.addEventListener("click", function () {
        self.direction = i < self.index ? "prev" : "next";
        self.index = i;
        self.renderSlides(true);
      });
    });
  };

  ImageCarousel.prototype.open = function () {
    if (this.isOpen) {
      return;
    }
    this.isOpen = true;
    this.lastFocused = document.activeElement;
    this.overlay.classList.add("is-open");
    this.overlay.removeAttribute("hidden");
    this.overlay.setAttribute("role", "dialog");
    this.overlay.setAttribute("aria-modal", "true");
    this.overlay.setAttribute(
      "aria-label",
      this.config.buttonText || "Image gallery"
    );
    document.body.style.overflow = "hidden";
    this.renderSlides(false);
    this.closeBtn.focus();
    document.addEventListener("keydown", this.handleKeydown);
    window.addEventListener("resize", this.handleResize);
  };

  ImageCarousel.prototype.close = function () {
    if (!this.isOpen) {
      return;
    }
    this.isOpen = false;
    this.overlay.classList.remove("is-open");
    this.overlay.setAttribute("hidden", "hidden");
    document.body.style.overflow = "";
    document.removeEventListener("keydown", this.handleKeydown);
    window.removeEventListener("resize", this.handleResize);
    if (this.lastFocused && typeof this.lastFocused.focus === "function") {
      this.lastFocused.focus();
    }
  };

  ImageCarousel.prototype.next = function () {
    if (this.items.length <= 1) {
      return;
    }
    this.direction = "next";
    this.index = (this.index + 1) % this.items.length;
    this.renderSlides(true);
  };

  ImageCarousel.prototype.prev = function () {
    if (this.items.length <= 1) {
      return;
    }
    this.direction = "prev";
    this.index = (this.index - 1 + this.items.length) % this.items.length;
    this.renderSlides(true);
  };

  ImageCarousel.prototype.handleKeydown = function (event) {
    if (!this.isOpen) {
      return;
    }
    if (event.key === "Escape") {
      this.close();
    } else if (event.key === "ArrowRight") {
      this.next();
    } else if (event.key === "ArrowLeft") {
      this.prev();
    }
  };

  ImageCarousel.prototype.handleResize = function () {
    if (this.isOpen) {
      this.renderSlides(false);
    }
  };

  function init() {
    var root = document.getElementById("itot-image-carousel");
    if (!root) {
      return;
    }
    new ImageCarousel(root, window.ITOT_IMAGE_CAROUSEL || {});
  }

  onReady(init);
})();
