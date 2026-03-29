import { IInputs, IOutputs } from "./generated/ManifestTypes";

// ─── Types ─────────────────────────────────────────────────────
interface RippleElement extends HTMLSpanElement {
  _timeout?: ReturnType<typeof setTimeout>;
}

// ─── GradientButton PCF Control ────────────────────────────────
export class GradientButton
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _container: HTMLDivElement;
  private _wrapper: HTMLDivElement;
  private _button: HTMLButtonElement;
  private _iconEl: HTMLSpanElement;
  private _labelEl: HTMLSpanElement;
  private _spinnerEl: HTMLSpanElement;
  private _notifyOutputChanged: () => void;
  private _context: ComponentFramework.Context<IInputs>;
  private _clickCount = 0;

  // ── Lifecycle: init ─────────────────────────────────────────
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._notifyOutputChanged = notifyOutputChanged;
    this._container = container;

    this._buildDOM();
    this._attachEvents();
    this._applyStyles();

    // Enable transitions after first paint to avoid FOUC
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        this._button.classList.add("transitions-ready");
      });
    });
  }

  // ── Lifecycle: updateView ────────────────────────────────────
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;
    this._applyStyles();
  }

  // ── Lifecycle: getOutputs ────────────────────────────────────
  public getOutputs(): IOutputs {
    return {
      onClickOutput: `click_${this._clickCount}`,
    };
  }

  // ── Lifecycle: destroy ───────────────────────────────────────
  public destroy(): void {
    this._button.removeEventListener("click", this._handleClick);
    this._button.removeEventListener("mouseenter", this._handleMouseEnter);
    this._button.removeEventListener("mouseleave", this._handleMouseLeave);
    this._button.removeEventListener("mousedown", this._handleMouseDown);
    this._button.removeEventListener("mouseup", this._handleMouseUp);
  }

  // ═══════════════════════════════════════════════════════════════
  // PRIVATE — DOM Construction
  // ═══════════════════════════════════════════════════════════════

  private _buildDOM(): void {
    // Wrapper
    this._wrapper = document.createElement("div");
    this._wrapper.className = "gradient-button-wrapper";

    // Button
    this._button = document.createElement("button");
    this._button.type = "button";
    this._button.className = "gradient-button";

    // Icon
    this._iconEl = document.createElement("span");
    this._iconEl.className = "gradient-button__icon";
    this._iconEl.setAttribute("aria-hidden", "true");

    // Label
    this._labelEl = document.createElement("span");
    this._labelEl.className = "gradient-button__label";

    // Spinner
    this._spinnerEl = document.createElement("span");
    this._spinnerEl.className = "gradient-button__spinner";
    this._spinnerEl.setAttribute("aria-hidden", "true");
    this._spinnerEl.style.display = "none";

    this._button.append(this._iconEl, this._spinnerEl, this._labelEl);
    this._wrapper.appendChild(this._button);
    this._container.appendChild(this._wrapper);
  }

  // ═══════════════════════════════════════════════════════════════
  // PRIVATE — Style Application
  // ═══════════════════════════════════════════════════════════════

  private _applyStyles(): void {
    const p = this._context.parameters;
    const btn = this._button;
    const wrapper = this._wrapper;

    // ── Disabled / Loading state ─────────────────────────────
    const isDisabled = p.isDisabled?.raw ?? false;
    const isLoading  = p.isLoading?.raw ?? false;
    const inactive   = isDisabled || isLoading;

    btn.disabled = inactive;
    btn.setAttribute("aria-disabled", String(inactive));

    // ── Label & tooltip ──────────────────────────────────────
    const label = p.label?.raw ?? "Click Me";
    const loadingText = p.loadingText?.raw ?? "Loading...";
    this._labelEl.textContent = isLoading ? loadingText : label;

    const tooltip = p.tooltipText?.raw ?? "";
    if (tooltip) btn.title = tooltip; else btn.removeAttribute("title");
    btn.setAttribute("aria-label", isLoading ? loadingText : label);

    // Spinner visibility
    this._spinnerEl.style.display = isLoading ? "inline-block" : "none";

    // ── Wrapper alignment ────────────────────────────────────
    const alignment = p.horizontalAlignment?.raw ?? "center";
    wrapper.className = `gradient-button-wrapper align-${alignment}`;

    // ── Build CSS custom properties on the button ────────────
    const style = btn.style;

    // Size
    style.setProperty("width",  this._sanitizeSize(p.buttonWidth?.raw ?? "auto"));
    style.setProperty("height", this._sanitizeSize(p.buttonHeight?.raw ?? "auto"));

    // Padding
    const pTop    = this._num(p.paddingTop?.raw, 12);
    const pBottom = this._num(p.paddingBottom?.raw, 12);
    const pLeft   = this._num(p.paddingLeft?.raw, 24);
    const pRight  = this._num(p.paddingRight?.raw, 24);
    style.setProperty("padding", `${pTop}px ${pRight}px ${pBottom}px ${pLeft}px`);

    // ── Typography ───────────────────────────────────────────
    const fontFamily  = p.fontFamily?.raw  ?? "Segoe UI, system-ui, sans-serif";
    const fontSize    = this._num(p.fontSize?.raw, 14);
    const fontWeight  = p.fontWeight?.raw  ?? "600";
    const fontStyle   = p.fontStyle?.raw   ?? "normal";
    const textColor   = inactive
      ? (p.disabledTextColor?.raw ?? "#666666")
      : (p.textColor?.raw ?? "#ffffff");
    const textTransform   = p.textTransform?.raw ?? "none";
    const letterSpacing   = this._num(p.letterSpacing?.raw, 0);
    const lineHeight      = this._num(p.lineHeight?.raw, 1.5);

    style.setProperty("font-family",     fontFamily);
    style.setProperty("font-size",       `${fontSize}px`);
    style.setProperty("font-weight",     fontWeight);
    style.setProperty("font-style",      fontStyle);
    style.setProperty("color",           textColor);
    style.setProperty("text-transform",  textTransform);
    style.setProperty("letter-spacing",  `${letterSpacing}em`);
    style.setProperty("line-height",     String(lineHeight));

    // Text shadow
    if (p.textShadowEnabled?.raw && !inactive) {
      const tsColor = p.textShadowColor?.raw ?? "rgba(0,0,0,0.3)";
      const tsBlur  = this._num(p.textShadowBlur?.raw, 4);
      style.setProperty("text-shadow", `0 1px ${tsBlur}px ${tsColor}`);
    } else {
      style.removeProperty("text-shadow");
    }

    // ── Border Radius ────────────────────────────────────────
    const br   = this._num(p.borderRadius?.raw, 8);
    const brTL = this._num(p.borderRadiusTopLeft?.raw, -1);
    const brTR = this._num(p.borderRadiusTopRight?.raw, -1);
    const brBL = this._num(p.borderRadiusBottomLeft?.raw, -1);
    const brBR = this._num(p.borderRadiusBottomRight?.raw, -1);

    const useFineControl = brTL >= 0 || brTR >= 0 || brBL >= 0 || brBR >= 0;
    if (useFineControl) {
      style.setProperty(
        "border-radius",
        `${brTL >= 0 ? brTL : br}px ${brTR >= 0 ? brTR : br}px ${brBR >= 0 ? brBR : br}px ${brBL >= 0 ? brBL : br}px`
      );
    } else {
      style.setProperty("border-radius", `${br}px`);
    }

    // ── Border ───────────────────────────────────────────────
    const gradientBorder = p.gradientBorder?.raw ?? false;
    const borderWidth    = this._num(p.borderWidth?.raw, 0);

    if (gradientBorder && borderWidth > 0 && !inactive) {
      const gbStart = p.gradientBorderColorStart?.raw ?? "#667eea";
      const gbEnd   = p.gradientBorderColorEnd?.raw   ?? "#764ba2";
      btn.classList.add("has-gradient-border");
      style.setProperty("--gb-border-width",    `${borderWidth}px`);
      style.setProperty("--gb-border-gradient", `linear-gradient(${this._num(p.gradientAngle?.raw, 135)}deg, ${gbStart}, ${gbEnd})`);
    } else {
      btn.classList.remove("has-gradient-border");
      const borderColor = p.borderColor?.raw  ?? "transparent";
      const borderStyle = p.borderStyle?.raw  ?? "solid";
      style.setProperty("border", `${borderWidth}px ${borderStyle} ${borderColor}`);
    }

    // ── Background (Gradient) ────────────────────────────────
    const bgGradient = this._buildBackground(p, inactive);
    style.setProperty("background", bgGradient);

    // ── Shadow & Glow ────────────────────────────────────────
    const shadowBox = this._buildShadow(p, inactive);
    style.setProperty("box-shadow", shadowBox || "none");
    style.setProperty("--gb-shadow", shadowBox || "none");

    // Glow for animations
    if (p.glowEnabled?.raw && !inactive) {
      const gc    = p.glowColor?.raw  ?? "rgba(102,126,234,0.6)";
      const gb    = this._num(p.glowBlur?.raw, 20);
      const gs    = this._num(p.glowSpread?.raw, 0);
      const glow  = `0 0 ${gb}px ${gs}px ${gc}`;
      const glowI = `0 0 ${gb * 2}px ${gs * 2}px ${gc}`;
      style.setProperty("--gb-glow",        `${glow}${shadowBox ? `, ${shadowBox}` : ""}`);
      style.setProperty("--gb-glow-intense", `${glowI}${shadowBox ? `, ${shadowBox}` : ""}`);
    }

    // Disabled opacity
    if (isDisabled) {
      style.setProperty("opacity", String(this._num(p.disabledOpacity?.raw, 0.6)));
    } else {
      style.removeProperty("opacity");
    }

    // ── Animation ────────────────────────────────────────────
    const anim     = p.animationType?.raw ?? "none";
    const animDur  = this._num(p.animationDuration?.raw, 2);
    const trans    = this._num(p.transitionDuration?.raw, 0.2);

    ["anim-shimmer", "anim-pulse", "anim-bounce", "anim-glow"].forEach(c => btn.classList.remove(c));
    if (anim !== "none" && !inactive) btn.classList.add(`anim-${anim}`);

    style.setProperty("--gb-anim-duration", `${animDur}s`);
    style.setProperty("--gb-transition",    `${trans}s`);

    // Hover scale / active scale stored as CSS vars (applied in event handlers)
    style.setProperty("--gb-hover-scale",  String(this._num(p.hoverScale?.raw, 1.03)));
    style.setProperty("--gb-active-scale", String(this._num(p.activeScale?.raw, 0.97)));
    style.setProperty("--gb-hover-brightness", String(this._num(p.hoverBrightness?.raw, 1.1)));

    // ── Hover gradient background (stored for event handler) ──
    if (p.hoverGradientColorStart?.raw && p.hoverGradientColorEnd?.raw) {
      const hAngle = this._num(p.hoverGradientAngle?.raw, 135);
      const hBg    = `linear-gradient(${hAngle}deg, ${p.hoverGradientColorStart.raw}, ${p.hoverGradientColorEnd.raw})`;
      style.setProperty("--gb-hover-bg", hBg);
      style.setProperty("--gb-default-bg", bgGradient);
    } else {
      style.removeProperty("--gb-hover-bg");
      style.removeProperty("--gb-default-bg");
    }

    // ── Icon ─────────────────────────────────────────────────
    this._applyIcon(p);
  }

  // ── Build background gradient string ─────────────────────────
  private _buildBackground(
    p: IInputs,
    inactive: boolean
  ): string {
    if (inactive) {
      const ds = p.disabledGradientColorStart?.raw ?? "#cccccc";
      const de = p.disabledGradientColorEnd?.raw   ?? "#aaaaaa";
      return `linear-gradient(135deg, ${ds}, ${de})`;
    }

    const gradType  = p.gradientType?.raw ?? "linear";
    const colorStart = p.gradientColorStart?.raw ?? "#667eea";
    const colorMid   = p.gradientColorMid?.raw   ?? "";
    const colorEnd   = p.gradientColorEnd?.raw   ?? "#764ba2";
    const angle      = this._num(p.gradientAngle?.raw, 135);
    const startPos   = this._num(p.gradientStartPosition?.raw, 0);
    const endPos     = this._num(p.gradientEndPosition?.raw, 100);

    const stops = colorMid
      ? `${colorStart} ${startPos}%, ${colorMid} 50%, ${colorEnd} ${endPos}%`
      : `${colorStart} ${startPos}%, ${colorEnd} ${endPos}%`;

    switch (gradType) {
      case "radial": return `radial-gradient(circle, ${stops})`;
      case "conic":  return `conic-gradient(from ${angle}deg, ${stops})`;
      case "solid":  return colorStart;
      default:       return `linear-gradient(${angle}deg, ${stops})`;
    }
  }

  // ── Build box-shadow string ───────────────────────────────────
  private _buildShadow(p: IInputs, inactive: boolean): string {
    const parts: string[] = [];

    // Main shadow
    if (p.shadowEnabled?.raw && !inactive) {
      const sc     = p.shadowColor?.raw  ?? "rgba(102,126,234,0.4)";
      const sx     = this._num(p.shadowOffsetX?.raw, 0);
      const sy     = this._num(p.shadowOffsetY?.raw, 4);
      const sb     = this._num(p.shadowBlur?.raw, 15);
      const ss     = this._num(p.shadowSpread?.raw, 0);
      const inset  = p.shadowInset?.raw ? "inset " : "";
      parts.push(`${inset}${sx}px ${sy}px ${sb}px ${ss}px ${sc}`);
    }

    // Glow
    if (p.glowEnabled?.raw && !inactive) {
      const gc  = p.glowColor?.raw  ?? "rgba(102,126,234,0.6)";
      const gb  = this._num(p.glowBlur?.raw, 20);
      const gs  = this._num(p.glowSpread?.raw, 0);
      parts.push(`0 0 ${gb}px ${gs}px ${gc}`);
    }

    return parts.join(", ");
  }

  // ── Apply icon from Fluent/Fabric icon name ───────────────────
  private _applyIcon(p: IInputs): void {
    const iconName = p.iconName?.raw ?? "";
    const iconPos  = p.iconPosition?.raw ?? "left";
    const iconSize = this._num(p.iconSize?.raw, 16);
    const iconGap  = this._num(p.iconGap?.raw, 8);
    const isLoading = p.isLoading?.raw ?? false;

    // Arrange icon + label flex direction
    ["icon-position-left", "icon-position-right",
     "icon-position-top",  "icon-position-bottom",
     "icon-position-only"].forEach(c => this._button.classList.remove(c));

    if (iconName && !isLoading) {
      this._iconEl.style.display    = "inline-flex";
      this._iconEl.style.fontSize   = `${iconSize}px`;
      this._iconEl.style.width      = `${iconSize}px`;
      this._iconEl.style.height     = `${iconSize}px`;
      this._iconEl.innerHTML        = this._getFluentIcon(iconName, iconSize);
      this._button.classList.add(`icon-position-${iconPos}`);

      const isVertical = iconPos === "top" || iconPos === "bottom";
      const gapProp    = isVertical ? "margin-bottom" : "margin-right";
      const revGapProp = isVertical ? "margin-top"    : "margin-left";

      this._iconEl.style.marginBottom = "0";
      this._iconEl.style.marginTop    = "0";
      this._iconEl.style.marginRight  = "0";
      this._iconEl.style.marginLeft   = "0";

      if (iconPos !== "only") {
        if (iconPos === "right" || iconPos === "bottom") {
          this._iconEl.style[revGapProp as "marginTop" | "marginLeft"] = `${iconGap}px`;
        } else {
          this._iconEl.style[gapProp as "marginBottom" | "marginRight"] = `${iconGap}px`;
        }
      }

      this._labelEl.style.display = iconPos === "only" ? "none" : "inline-flex";
    } else {
      this._iconEl.style.display  = "none";
      this._iconEl.innerHTML      = "";
      this._labelEl.style.display = "inline-flex";
    }

    // Loading spinner gap
    if (isLoading) {
      this._spinnerEl.style.marginRight = `${iconGap}px`;
      this._spinnerEl.style.fontSize    = `${iconSize}px`;
    }
  }

  // ── SVG icon resolver (Fluent UI system icons) ────────────────
  private _getFluentIcon(name: string, size: number): string {
    const icons: Record<string, string> = {
      save:    `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M17 3H5a2 2 0 00-2 2v14a2 2 0 002 2h14a2 2 0 002-2V7l-4-4zm-5 16a3 3 0 110-6 3 3 0 010 6zm3-10H5V5h10v4z"/></svg>`,
      send:    `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M2.01 21L23 12 2.01 3 2 10l15 2-15 2z"/></svg>`,
      add:     `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>`,
      delete:  `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M6 19a2 2 0 002 2h8a2 2 0 002-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>`,
      edit:    `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04a1 1 0 000-1.41l-2.34-2.34a1 1 0 00-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>`,
      search:  `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M15.5 14h-.79l-.28-.27A6.47 6.47 0 0016 9.5 6.5 6.5 0 109.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>`,
      check:   `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>`,
      close:   `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/></svg>`,
      download:`<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>`,
      upload:  `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M9 16h6v-6h4l-7-7-7 7h4v6zm-4 2h14v2H5v-2z"/></svg>`,
      refresh: `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M17.65 6.35A7.958 7.958 0 0012 4a8 8 0 00-8 8 8 8 0 008 8c3.73 0 6.84-2.55 7.73-6h-2.08A5.99 5.99 0 0112 18a6 6 0 01-6-6 6 6 0 016-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z"/></svg>`,
      filter:  `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M10 18h4v-2h-4v2zM3 6v2h18V6H3zm3 7h12v-2H6v2z"/></svg>`,
      print:   `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M19 8H5a3 3 0 00-3 3v6h4v4h12v-4h4v-6a3 3 0 00-3-3zm-3 11H8v-5h8v5zm3-7a1 1 0 110-2 1 1 0 010 2zm-1-9H6v4h12V3z"/></svg>`,
      export:  `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M9 3L5 6.99h3V14h2V6.99h3L9 3zm7 14.01V10h-2v7.01h-3L15 21l4-3.99h-3z"/></svg>`,
      arrow_right: `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>`,
    };

    return icons[name.toLowerCase()] ?? `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="currentColor"><circle cx="12" cy="12" r="10" fill-opacity=".3"/><text x="12" y="16" text-anchor="middle" font-size="10">${name.charAt(0).toUpperCase()}</text></svg>`;
  }

  // ═══════════════════════════════════════════════════════════════
  // PRIVATE — Event Handlers
  // ═══════════════════════════════════════════════════════════════

  private _handleClick = (e: MouseEvent): void => {
    if (this._button.disabled) return;

    // Ripple
    if (this._context.parameters.rippleEffect?.raw ?? true) {
      this._createRipple(e);
    }

    this._clickCount++;
    this._notifyOutputChanged();
  };

  private _handleMouseEnter = (): void => {
    if (this._button.disabled) return;
    const s = this._button.style;
    const scale      = s.getPropertyValue("--gb-hover-scale")      || "1.03";
    const brightness = s.getPropertyValue("--gb-hover-brightness") || "1.1";
    const hoverBg    = s.getPropertyValue("--gb-hover-bg");

    s.setProperty("transform", `scale(${scale})`);
    s.setProperty("filter", `brightness(${brightness})`);
    if (hoverBg) s.setProperty("background", hoverBg);
  };

  private _handleMouseLeave = (): void => {
    if (this._button.disabled) return;
    const s = this._button.style;
    s.removeProperty("transform");
    s.removeProperty("filter");
    const defaultBg = s.getPropertyValue("--gb-default-bg");
    if (defaultBg) s.setProperty("background", defaultBg);
    this._handleMouseUp();
  };

  private _handleMouseDown = (): void => {
    if (this._button.disabled) return;
    const scale = this._button.style.getPropertyValue("--gb-active-scale") || "0.97";
    this._button.style.setProperty("transform", `scale(${scale})`);
  };

  private _handleMouseUp = (): void => {
    if (this._button.disabled) return;
    const scale = this._button.style.getPropertyValue("--gb-hover-scale") || "1.03";
    this._button.style.setProperty("transform", `scale(${scale})`);
  };

  private _attachEvents(): void {
    this._button.addEventListener("click",      this._handleClick);
    this._button.addEventListener("mouseenter", this._handleMouseEnter);
    this._button.addEventListener("mouseleave", this._handleMouseLeave);
    this._button.addEventListener("mousedown",  this._handleMouseDown);
    this._button.addEventListener("mouseup",    this._handleMouseUp);
  }

  // ── Ripple effect ─────────────────────────────────────────────
  private _createRipple(e: MouseEvent): void {
    const btn    = this._button;
    const rect   = btn.getBoundingClientRect();
    const size   = Math.max(rect.width, rect.height) * 2;
    const x      = e.clientX - rect.left - size / 2;
    const y      = e.clientY - rect.top  - size / 2;
    const color  = this._context.parameters.rippleColor?.raw ?? "rgba(255,255,255,0.3)";

    const ripple = document.createElement("span") as RippleElement;
    ripple.className = "gradient-button__ripple";
    Object.assign(ripple.style, {
      width:           `${size}px`,
      height:          `${size}px`,
      left:            `${x}px`,
      top:             `${y}px`,
      backgroundColor: color,
    });

    btn.appendChild(ripple);
    ripple._timeout = setTimeout(() => ripple.remove(), 650);
  }

  // ═══════════════════════════════════════════════════════════════
  // PRIVATE — Helpers
  // ═══════════════════════════════════════════════════════════════

  private _num(value: number | null | undefined, fallback: number): number {
    return value !== null && value !== undefined && !isNaN(Number(value))
      ? Number(value)
      : fallback;
  }

  private _sanitizeSize(value: string): string {
    if (!value || value === "auto") return "auto";
    if (/^\d+$/.test(value)) return `${value}px`;
    return value;
  }
}
