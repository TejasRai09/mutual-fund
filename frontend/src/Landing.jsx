import { useState, useEffect, useRef } from 'react'
import { ArrowRight, CheckCircle2, FileSpreadsheet, Zap, Download, BarChart3, Lock, Brain } from 'lucide-react'

// ── Brand tokens ────────────────────────────────────────────────────────────
const NAVY      = '#1B3080'
const NAVY_DARK = '#0F1F5C'
const RED       = '#E5312C'
const GREEN     = '#3BAE35'
const TEXT      = '#0f172a'
const MUTED     = '#64748b'
const BORDER    = '#e2e8f0'
const BG_ALT    = '#f8fafc'

const ZUARI_LOGO   = 'https://www.zuarimoney.com/App_Themes/images/Zuari_whatsnew2.jpg'
const ADVENTZ_LOGO = 'https://www.adventz.com/html/images/logo.png'

// ── Static data ─────────────────────────────────────────────────────────────
const AMC_LIST = [
  'HDFC', 'SBI', 'ICICI Prudential', 'Nippon India', 'Aditya Birla Sun Life',
  'Axis', 'DSP', 'Mirae Asset', 'Franklin Templeton', 'PGIM India',
  'Bandhan', 'TATA', 'Motilal Oswal', 'Invesco', 'Canara Robeco',
  'Sundaram', 'HSBC', 'Bank of India', 'LIC MF', 'Mahindra Manulife',
]

const STEPS = [
  {
    num: '01', title: 'Upload Source PDFs',
    desc: 'Drop AMC brokerage rate PDFs and your Excel templates. Supports 19+ AMC formats. Enter password once for all encrypted files.',
    bullets: ['19+ AMC PDF formats', 'AES-encrypted PDFs supported', 'Multiple files in one batch'],
    preview: [
      { name: 'HDFC_Brokerage_Apr.pdf',  status: '31 schemes' },
      { name: 'SBI_Trail_Rates.pdf',     status: '24 schemes' },
      { name: 'AXIS_Brokerage.pdf',      status: '22 schemes' },
      { name: 'Target_Template.xlsx',    status: 'Ready' },
    ],
    previewLabel: 'FILES UPLOADED',
  },
  {
    num: '02', title: 'AI Matches Schemes',
    desc: 'Fuzzy matching handles typos, abbreviations, old AMC names, and suffix variants. Leaves blanks instead of wrong values.',
    bullets: ['"Reliance MF" → Nippon India', 'Typo-tolerant normalisation', 'BPS and decimal conversion'],
    preview: [
      { name: 'Flexi Cap Fund',   status: '0.65%' },
      { name: 'Midcap 150 Fund',  status: '0.70%' },
      { name: 'ELSS Tax Saver',   status: '0.60%' },
      { name: 'Small Cap Fund',   status: '0.75%' },
    ],
    previewLabel: 'MATCHED SCHEMES (EX-GST)',
  },
  {
    num: '03', title: 'Download Filled Excel',
    desc: 'All templates filled with EX-GST trail values and bundled into a ZIP. 4 trail columns per scheme. Ready to submit.',
    bullets: ['EX-GST values only — never IN-GST', '4 trail columns per scheme', 'Instant ZIP download'],
    preview: [
      { name: 'HDFC_Template_Filled.xlsx',  status: '31/32' },
      { name: 'SBI_Template_Filled.xlsx',   status: '24/24' },
      { name: 'AXIS_Template_Filled.xlsx',  status: '22/22' },
    ],
    previewLabel: 'FILES READY',
  },
]

const FEATURES = [
  { icon: FileSpreadsheet, title: 'Smart PDF Parsing',     desc: 'Extracts rates from tables, paragraphs, and multi-page AMC documents. Password-protected files handled with a single password entry.' },
  { icon: Brain,           title: 'Fuzzy Scheme Matching', desc: 'Multi-step algorithm: exact → normalised → partial → word-subset. Handles old names ("Reliance" → Nippon), abbreviations, and typos.' },
  { icon: Zap,             title: 'EX-GST Accuracy',       desc: 'All trail values written in EX-GST form. Handles BPS conversion (Motilal), decimal fractions (ICICI), and varied trail structures.' },
  { icon: Download,        title: 'Batch Processing',       desc: 'Upload all AMC PDFs and templates at once. One click fills every file and bundles them into a single download ZIP.' },
  { icon: Lock,            title: 'Encrypted PDF Support',  desc: 'AXIS, BOI and other protected files work seamlessly. One password entry unlocks all encrypted files in the batch.' },
  { icon: BarChart3,       title: 'Multi-format Excel',     desc: 'Auto-detects column-based Redos and row-based Vijayinfotech templates. No manual format selection needed.' },
]

// ── Hooks ───────────────────────────────────────────────────────────────────
function useReveal() {
  useEffect(() => {
    const obs = new IntersectionObserver(
      entries => entries.forEach(e => {
        if (e.isIntersecting) { e.target.classList.add('visible'); obs.unobserve(e.target) }
      }),
      { threshold: 0.08 }
    )
    document.querySelectorAll('.reveal').forEach(el => obs.observe(el))
    return () => obs.disconnect()
  }, [])
}

function Counter({ target, suffix = '' }) {
  const [val, setVal]         = useState(0)
  const [started, setStarted] = useState(false)
  const ref                   = useRef(null)

  useEffect(() => {
    const el = ref.current; if (!el) return
    const obs = new IntersectionObserver(
      ([e]) => { if (e.isIntersecting && !started) setStarted(true) },
      { threshold: 0.5 }
    )
    obs.observe(el); return () => obs.disconnect()
  }, [started])

  useEffect(() => {
    if (!started) return
    const dur = 1600, t0 = Date.now()
    const tick = () => {
      const p = Math.min((Date.now() - t0) / dur, 1)
      setVal(Math.round((1 - Math.pow(1 - p, 3)) * target))
      if (p < 1) requestAnimationFrame(tick)
    }
    requestAnimationFrame(tick)
  }, [started, target])

  return <span ref={ref}>{val}{suffix}</span>
}

// ── Landing ─────────────────────────────────────────────────────────────────
export default function Landing({ onLaunch }) {
  const [activeStep, setActiveStep] = useState(0)
  useReveal()

  const amcDoubled = [...AMC_LIST, ...AMC_LIST]

  return (
    <div style={{ background: '#fff', color: TEXT, minHeight: '100vh', overflowX: 'hidden' }}>

      {/* ── Group top-bar ───────────────────────────────────────────────── */}
      <div style={{
        background: NAVY_DARK, padding: '7px 60px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
      }}>
        <span style={{ color: 'rgba(255,255,255,0.45)', fontSize: 11, letterSpacing: '0.06em' }}>
          Zuari Finserv — Mutual Fund Brokerage Automation
        </span>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span style={{ color: 'rgba(255,255,255,0.45)', fontSize: 11 }}>An Adventz Group Company</span>
          {/* White badge so Adventz logo renders in its original colours */}
          <div style={{
            background: '#fff', borderRadius: 4, padding: '3px 8px',
            display: 'flex', alignItems: 'center',
          }}>
            <img src={ADVENTZ_LOGO} alt="Adventz" style={{ height: 18, objectFit: 'contain', display: 'block' }} />
          </div>
        </div>
      </div>

      {/* ── Navbar ──────────────────────────────────────────────────────── */}
      <nav style={{
        position: 'sticky', top: 0, zIndex: 100,
        background: '#fff',
        borderBottom: `3px solid ${RED}`,
        padding: '0 60px', height: 82,
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        boxShadow: '0 2px 16px rgba(0,0,0,0.08)',
      }}>
        {/* Logo + divider + tagline */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 18 }}>
          <img src={ZUARI_LOGO} alt="Zuari Finserv" style={{ height: 58, objectFit: 'contain' }} />
          <div style={{ width: 1, height: 36, background: BORDER }} />
          <span style={{ fontSize: 12, color: MUTED, fontWeight: 500, letterSpacing: '0.02em' }}>
            Brokerage Automation Tool
          </span>
        </div>

        {/* Nav links + CTA */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 36 }}>
          <NavLink href="#process">How it Works</NavLink>
          <NavLink href="#amcs">Supported AMCs</NavLink>
          <NavLink href="#features">Features</NavLink>
          {/* Red CTA — matches the red in the Zuari Z */}
          <button onClick={onLaunch} style={{
            display: 'inline-flex', alignItems: 'center', gap: 8,
            padding: '11px 26px', borderRadius: 6,
            background: `linear-gradient(135deg, ${RED} 0%, #C02020 100%)`,
            color: '#fff', fontSize: 13, fontWeight: 700,
            border: 'none', cursor: 'pointer',
            boxShadow: '0 4px 14px rgba(229,49,44,0.38)',
            transition: 'all 0.15s',
            letterSpacing: '0.01em',
          }}
            onMouseEnter={e => { e.currentTarget.style.transform = 'translateY(-1px)'; e.currentTarget.style.boxShadow = '0 6px 20px rgba(229,49,44,0.5)' }}
            onMouseLeave={e => { e.currentTarget.style.transform = 'translateY(0)'; e.currentTarget.style.boxShadow = '0 4px 14px rgba(229,49,44,0.38)' }}
          >
            Launch Tool <ArrowRight size={14} />
          </button>
        </div>
      </nav>

      {/* ── Hero ────────────────────────────────────────────────────────── */}
      <section style={{ background: BG_ALT, borderBottom: `1px solid ${BORDER}` }}>
        <div style={{
          maxWidth: 1200, margin: '0 auto', padding: '72px 48px 80px',
          display: 'grid', gridTemplateColumns: '1fr 370px', gap: 64, alignItems: 'center',
        }}>

          {/* Left */}
          <div>
            <div className="anim-slide-up" style={{
              display: 'inline-flex', alignItems: 'center', gap: 8,
              padding: '4px 12px', borderRadius: 4,
              background: `${NAVY}10`, color: NAVY,
              fontSize: 11, fontWeight: 700, letterSpacing: '0.1em',
              border: `1px solid ${NAVY}22`, marginBottom: 28,
            }}>
              AI-POWERED BROKERAGE AUTOMATION
            </div>

            <h1 className="anim-slide-up" style={{
              fontSize: 'clamp(40px, 5vw, 68px)',
              fontWeight: 900, lineHeight: 1.0,
              letterSpacing: '-0.03em', margin: '0 0 22px',
              animationDelay: '0.08s',
            }}>
              Automate AMC<br />Brokerage Filing.
            </h1>

            <p className="anim-slide-up" style={{
              fontSize: 16, lineHeight: 1.75, color: MUTED,
              maxWidth: 480, margin: '0 0 36px',
              animationDelay: '0.14s',
            }}>
              Upload brokerage PDFs from 19+ AMCs, drop your Excel templates —
              the tool matches every scheme and fills EX-GST trail rates.
              No manual work. No wrong values.
            </p>

            <div className="anim-slide-up" style={{
              display: 'flex', gap: 12, alignItems: 'center',
              marginBottom: 32, animationDelay: '0.2s',
            }}>
              <PrimaryBtn onClick={onLaunch}>
                Launch Tool <ArrowRight size={15} />
              </PrimaryBtn>
              <GhostBtn href="#process">How it works</GhostBtn>
            </div>

            <div className="anim-slide-up" style={{
              display: 'flex', gap: 20, flexWrap: 'wrap',
              animationDelay: '0.26s',
            }}>
              {['19+ AMC PDFs', 'EX-GST values only', 'Fuzzy scheme matching', 'Encrypted PDF support'].map(t => (
                <div key={t} style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  <CheckCircle2 size={13} color={GREEN} />
                  <span style={{ fontSize: 12, color: MUTED }}>{t}</span>
                </div>
              ))}
            </div>
          </div>

          {/* Right: mock results card */}
          <div className="anim-slide-up" style={{
            background: '#fff', borderRadius: 10,
            border: `1px solid ${BORDER}`,
            boxShadow: '0 4px 24px rgba(0,0,0,0.07)',
            overflow: 'hidden',
            animationDelay: '0.3s',
          }}>
            <div style={{
              padding: '14px 20px', background: NAVY,
              display: 'flex', alignItems: 'center', justifyContent: 'space-between',
            }}>
              <span style={{ fontSize: 11, fontWeight: 700, color: 'rgba(255,255,255,0.7)', letterSpacing: '0.1em' }}>
                PROCESSING RESULTS
              </span>
              <div style={{ width: 6, height: 6, borderRadius: '50%', background: GREEN }} />
            </div>

            <div style={{ padding: '16px 20px' }}>
              {[
                { amc: 'HDFC',    filled: '31/32', trail: '0.65%' },
                { amc: 'SBI',     filled: '24/24', trail: '0.70%' },
                { amc: 'AXIS',    filled: '22/22', trail: '0.60%' },
                { amc: 'NIPPON',  filled: '28/29', trail: '0.75%' },
                { amc: 'MOTILAL', filled: '18/18', trail: '0.55%' },
              ].map(({ amc, filled, trail }) => (
                <div key={amc} style={{
                  display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                  padding: '9px 0', borderBottom: `1px solid ${BORDER}`,
                }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <div style={{ width: 5, height: 5, borderRadius: '50%', background: GREEN }} />
                    <span style={{ fontSize: 12, fontWeight: 600 }}>{amc}</span>
                  </div>
                  <div>
                    <span style={{ fontSize: 12, fontWeight: 700, color: NAVY }}>{filled}</span>
                    <span style={{ fontSize: 11, color: MUTED, marginLeft: 8 }}>{trail}</span>
                  </div>
                </div>
              ))}

              <div style={{
                marginTop: 12, padding: '10px 14px',
                background: `${NAVY}08`, border: `1px solid ${NAVY}18`,
                borderRadius: 6, display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              }}>
                <span style={{ fontSize: 12, color: MUTED }}>Total matched</span>
                <span style={{ fontSize: 15, fontWeight: 900, color: NAVY }}>123/125</span>
              </div>

              <button onClick={onLaunch} style={{
                width: '100%', marginTop: 10, padding: '10px 0',
                borderRadius: 6, background: NAVY, color: '#fff',
                fontSize: 13, fontWeight: 700, cursor: 'pointer', border: 'none',
              }}>
                Download ZIP →
              </button>
            </div>
          </div>
        </div>
      </section>

      {/* ── Stats strip ─────────────────────────────────────────────────── */}
      <section style={{ background: NAVY }}>
        <div style={{
          maxWidth: 1100, margin: '0 auto', padding: '36px 48px',
          display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)',
        }}>
          {[
            { target: 19,  suf: '+', label: 'AMCs Supported'    },
            { target: 45,  suf: '×', label: 'Faster Matching'   },
            { target: 5,   suf: 's', label: 'Avg Process Time'  },
            { target: 350, suf: '+', label: 'Schemes Per Batch' },
          ].map(({ target, suf, label }, i) => (
            <div key={label} style={{
              textAlign: 'center',
              borderLeft: i > 0 ? '1px solid rgba(255,255,255,0.12)' : 'none',
              padding: '4px 24px',
            }}>
              <div style={{
                fontSize: 'clamp(32px, 4vw, 52px)',
                fontWeight: 900, color: '#fff',
                letterSpacing: '-0.03em', lineHeight: 1,
              }}>
                <Counter target={target} suffix={suf} />
              </div>
              <div style={{ fontSize: 12, color: 'rgba(255,255,255,0.5)', marginTop: 6, fontWeight: 500 }}>
                {label}
              </div>
            </div>
          ))}
        </div>
      </section>

      {/* ── AMC ticker ──────────────────────────────────────────────────── */}
      <section id="amcs" style={{ borderTop: `1px solid ${BORDER}`, borderBottom: `1px solid ${BORDER}`, overflow: 'hidden' }}>
        <div style={{ padding: '12px 0', overflow: 'hidden' }}>
          <div className="marquee-track" style={{ display: 'flex', width: 'max-content' }}>
            {amcDoubled.map((amc, i) => (
              <div key={i} style={{
                padding: '0 28px', borderRight: `1px solid ${BORDER}`,
                fontSize: 11, fontWeight: 700, letterSpacing: '0.12em', textTransform: 'uppercase',
                color: i % 5 === 0 ? NAVY : MUTED,
                whiteSpace: 'nowrap',
              }}>
                {amc}
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* ── How it works ────────────────────────────────────────────────── */}
      <section id="process" style={{ padding: '80px 48px', background: BG_ALT, borderBottom: `1px solid ${BORDER}` }}>
        <div style={{ maxWidth: 1100, margin: '0 auto' }}>

          <div className="reveal" style={{ marginBottom: 52 }}>
            <div style={{ fontSize: 11, color: RED, fontWeight: 700, letterSpacing: '0.14em', marginBottom: 12 }}>
              HOW IT WORKS
            </div>
            <h2 style={{
              fontSize: 'clamp(28px, 3.5vw, 48px)', fontWeight: 900,
              letterSpacing: '-0.03em', margin: '0 0 10px',
            }}>
              Three steps. Zero manual work.
            </h2>
            <p style={{ color: MUTED, fontSize: 15, maxWidth: 440, margin: 0 }}>
              From raw AMC PDFs to filled Excel templates in under a minute.
            </p>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: '230px 1fr', gap: 28, alignItems: 'start' }}>

            {/* Step tabs */}
            <div className="reveal" style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              {STEPS.map((step, i) => (
                <button key={step.num} onClick={() => setActiveStep(i)} style={{
                  textAlign: 'left', padding: '14px 16px', borderRadius: 6,
                  cursor: 'pointer', transition: 'all 0.15s', border: 'none',
                  borderLeft: `4px solid ${activeStep === i ? NAVY : 'transparent'}`,
                  background: activeStep === i ? '#fff' : 'transparent',
                  boxShadow: activeStep === i ? '0 1px 6px rgba(0,0,0,0.06)' : 'none',
                }}>
                  <div style={{
                    fontSize: 10, fontWeight: 700, letterSpacing: '0.1em',
                    color: activeStep === i ? RED : MUTED, marginBottom: 3,
                  }}>
                    STEP {step.num}
                  </div>
                  <div style={{
                    fontSize: 13, fontWeight: 600,
                    color: activeStep === i ? TEXT : MUTED,
                  }}>
                    {step.title}
                  </div>
                </button>
              ))}
            </div>

            {/* Tab content */}
            <div className="reveal" style={{
              padding: '28px 28px', borderRadius: 8,
              border: `1px solid ${BORDER}`, background: '#fff',
              display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 28,
              boxShadow: '0 2px 12px rgba(0,0,0,0.04)',
            }}>
              <div>
                <div style={{
                  fontSize: 10, color: RED, fontWeight: 700,
                  letterSpacing: '0.1em', marginBottom: 10,
                }}>
                  STEP {STEPS[activeStep].num}
                </div>
                <h3 style={{
                  fontSize: 22, fontWeight: 800, margin: '0 0 12px',
                  letterSpacing: '-0.02em',
                }}>
                  {STEPS[activeStep].title}
                </h3>
                <p style={{ color: MUTED, fontSize: 14, lineHeight: 1.72, margin: '0 0 20px' }}>
                  {STEPS[activeStep].desc}
                </p>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 9 }}>
                  {STEPS[activeStep].bullets.map(b => (
                    <div key={b} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                      <CheckCircle2 size={14} color={GREEN} />
                      <span style={{ fontSize: 13, color: TEXT }}>{b}</span>
                    </div>
                  ))}
                </div>
              </div>

              {/* Mock preview */}
              <div style={{
                background: BG_ALT, borderRadius: 6,
                border: `1px solid ${BORDER}`, padding: '14px 16px',
              }}>
                <div style={{
                  fontSize: 10, color: MUTED, fontWeight: 700, letterSpacing: '0.1em',
                  marginBottom: 10, paddingBottom: 8, borderBottom: `1px solid ${BORDER}`,
                }}>
                  {STEPS[activeStep].previewLabel}
                </div>
                {STEPS[activeStep].preview.map((item, i, arr) => (
                  <div key={i} style={{
                    display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                    padding: '8px 0',
                    borderBottom: i < arr.length - 1 ? `1px solid ${BORDER}` : 'none',
                  }}>
                    <span style={{
                      fontSize: 11, color: MUTED,
                      maxWidth: '58%', overflow: 'hidden',
                      textOverflow: 'ellipsis', whiteSpace: 'nowrap',
                    }}>
                      {item.name}
                    </span>
                    <span style={{ fontSize: 11, fontWeight: 700, color: NAVY }}>
                      {item.status}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* ── Features ────────────────────────────────────────────────────── */}
      <section id="features" style={{ padding: '80px 48px', borderBottom: `1px solid ${BORDER}` }}>
        <div style={{ maxWidth: 1100, margin: '0 auto' }}>

          <div className="reveal" style={{ marginBottom: 48 }}>
            <div style={{ fontSize: 11, color: RED, fontWeight: 700, letterSpacing: '0.14em', marginBottom: 12 }}>
              CAPABILITIES
            </div>
            <h2 style={{
              fontSize: 'clamp(28px, 3.5vw, 48px)', fontWeight: 900,
              letterSpacing: '-0.03em', margin: '0 0 10px',
            }}>
              Everything you need.
            </h2>
            <p style={{ color: MUTED, fontSize: 15, maxWidth: 440, margin: 0 }}>
              Built specifically for Indian mutual fund brokerage workflows.
            </p>
          </div>

          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 20 }}>
            {FEATURES.map(({ icon: Icon, title, desc }, i) => (
              <FeatureCard key={title} Icon={Icon} title={title} desc={desc} delay={i * 60} />
            ))}
          </div>
        </div>
      </section>

      {/* ── CTA ─────────────────────────────────────────────────────────── */}
      <section style={{ background: NAVY, padding: '72px 48px' }}>
        <div className="reveal" style={{ maxWidth: 640, margin: '0 auto', textAlign: 'center' }}>
          <div style={{
            fontSize: 11, color: 'rgba(255,255,255,0.45)',
            fontWeight: 700, letterSpacing: '0.14em', marginBottom: 16,
          }}>
            GET STARTED
          </div>
          <h2 style={{
            fontSize: 'clamp(26px, 3.5vw, 44px)',
            fontWeight: 900, letterSpacing: '-0.03em',
            color: '#fff', margin: '0 0 16px', lineHeight: 1.05,
          }}>
            Stop filling brokerage slabs manually.
          </h2>
          <p style={{
            color: 'rgba(255,255,255,0.5)',
            maxWidth: 400, margin: '0 auto 32px',
            fontSize: 15, lineHeight: 1.72,
          }}>
            19+ AMC PDFs × 30+ scheme slabs = hours of monthly work.
            This tool eliminates it completely.
          </p>
          <button onClick={onLaunch} style={{
            display: 'inline-flex', alignItems: 'center', gap: 8,
            padding: '13px 28px', borderRadius: 6,
            background: '#fff', color: NAVY,
            fontSize: 14, fontWeight: 700, cursor: 'pointer', border: 'none',
            transition: 'all 0.15s',
          }}
            onMouseEnter={e => { e.currentTarget.style.background = '#eef2ff'; e.currentTarget.style.transform = 'translateY(-1px)' }}
            onMouseLeave={e => { e.currentTarget.style.background = '#fff'; e.currentTarget.style.transform = 'translateY(0)' }}
          >
            Launch Tool <ArrowRight size={15} />
          </button>
        </div>
      </section>

      {/* ── Footer ──────────────────────────────────────────────────────── */}
      <footer style={{
        background: BG_ALT, borderTop: `1px solid ${BORDER}`,
        padding: '24px 48px',
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        flexWrap: 'wrap', gap: 12,
      }}>
        <img src={ZUARI_LOGO} alt="Zuari Finserv" style={{ height: 36, objectFit: 'contain' }} />
        <span style={{ color: MUTED, fontSize: 12 }}>
          EX-GST · 19+ AMCs · Fuzzy matching · Encrypted PDF support
        </span>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ color: MUTED, fontSize: 11 }}>An Adventz Group Company</span>
          <img src={ADVENTZ_LOGO} alt="Adventz" style={{ height: 20, objectFit: 'contain' }} />
        </div>
      </footer>
    </div>
  )
}

// ── Sub-components ───────────────────────────────────────────────────────────

function NavLink({ href, children }) {
  const [hov, setHov] = useState(false)
  return (
    <a href={href} style={{
      color: hov ? NAVY : TEXT,
      fontSize: 13, fontWeight: 500, textDecoration: 'none', transition: 'color 0.15s',
    }} onMouseEnter={() => setHov(true)} onMouseLeave={() => setHov(false)}>
      {children}
    </a>
  )
}

function PrimaryBtn({ onClick, children }) {
  const [hov, setHov] = useState(false)
  return (
    <button onClick={onClick} style={{
      display: 'inline-flex', alignItems: 'center', gap: 7,
      padding: '12px 26px', borderRadius: 6,
      background: `linear-gradient(135deg, ${RED} 0%, #C02020 100%)`,
      color: '#fff', fontSize: 14, fontWeight: 700,
      border: 'none', cursor: 'pointer',
      transform: hov ? 'translateY(-2px)' : 'translateY(0)',
      boxShadow: hov ? '0 6px 20px rgba(229,49,44,0.5)' : '0 3px 12px rgba(229,49,44,0.3)',
      transition: 'all 0.15s',
    }}
      onMouseEnter={() => setHov(true)}
      onMouseLeave={() => setHov(false)}
    >
      {children}
    </button>
  )
}

function GhostBtn({ href, children }) {
  const [hov, setHov] = useState(false)
  return (
    <a href={href} style={{
      display: 'inline-flex', alignItems: 'center', gap: 6,
      padding: '10px 20px', borderRadius: 6,
      border: `1px solid ${hov ? NAVY : BORDER}`,
      background: '#fff', color: hov ? NAVY : TEXT,
      fontSize: 13, fontWeight: 600, textDecoration: 'none', transition: 'all 0.15s',
    }}
      onMouseEnter={() => setHov(true)}
      onMouseLeave={() => setHov(false)}
    >
      {children}
    </a>
  )
}

function FeatureCard({ Icon, title, desc, delay }) {
  const [hov, setHov] = useState(false)
  return (
    <div className="reveal" style={{
      padding: '24px 22px', borderRadius: 8, cursor: 'default',
      background: '#fff',
      border: `1px solid ${BORDER}`,
      borderTop: `3px solid ${hov ? NAVY : BORDER}`,
      boxShadow: hov ? '0 4px 16px rgba(27,48,128,0.09)' : '0 1px 4px rgba(0,0,0,0.03)',
      transitionDelay: `${delay}ms`,
      transition: 'border-color 0.2s, box-shadow 0.2s',
    }}
      onMouseEnter={() => setHov(true)}
      onMouseLeave={() => setHov(false)}
    >
      <div style={{
        width: 36, height: 36, borderRadius: 8, marginBottom: 14,
        background: hov ? `${NAVY}10` : BG_ALT,
        border: `1px solid ${hov ? NAVY + '28' : BORDER}`,
        display: 'flex', alignItems: 'center', justifyContent: 'center',
        transition: 'all 0.2s',
      }}>
        <Icon size={16} color={hov ? NAVY : MUTED} />
      </div>
      <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 8, color: TEXT }}>{title}</div>
      <div style={{ fontSize: 12, color: MUTED, lineHeight: 1.65 }}>{desc}</div>
    </div>
  )
}
