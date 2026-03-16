import { LightningElement, api, track } from 'lwc';
import { loadScript } from 'lightning/platformResourceLoader';
import D3 from '@salesforce/resourceUrl/d3';

/* ---------------- CONSTANTS ---------------- */

const STOP_WORDS = new Set([
    'a','an','and','are','as','at','be','by','for','from','has','he','in','is','it',
    'its','of','on','that','the','to','was','will','with','this','but','they','have',
    'had','what','when','where','who','which','why','how','or','not','so','than',
    'too','very','can','just','should','now','i','you','we','me','my','your','our',
    'their','his','her','am','been','being','do','does','did','doing','would',
    'could','about','after','again','against','all','also','any','because','before',
    'between','both','down','during','each','few','further','here','him','into',
    'more','most','must','no','nor','off','once','only','other','out','over','own',
    'same','she','some','such','then','there','these','those','through','under',
    'until','up','us','were','while','whom','yes','yet'
]);

const COLOR_PALETTE = [
    '#667eea','#764ba2','#f093fb','#4facfe','#43e97b',
    '#fa709a','#fee140','#30cfd0','#a8edea','#ff6b6b'
];

const ORIENTATION = {
    HORIZONTAL: 'horizontal',
    VERTICAL: 'vertical',
    BOTH: 'both'
};

const CHART_TYPES = {
    WORD_CLOUD: 'wordCloud',
    HORIZONTAL_WORD_CLOUD: 'h-wordCloud',
    VERTICAL_WORD_CLOUD: 'v-wordCloud',
    MIXED_WORD_CLOUD: 'm-wordCloud',
    HORIZONTAL_BAR: 'h-BarChart',
    VERTICAL_BAR: 'v-BarChart'
};

/* ---------------- COMPONENT ---------------- */

export default class TextWordCloudChart extends LightningElement {

    /* ---------- API INPUT ---------- */

    _results;

    @api
    get results() {
        return this._results;
    }

    set results(value) {
        this._results = value;
        console.log('Results set:', value ? value.length : 0, 'items');

        if (this.d3Initialized && Array.isArray(value) && value.length > 0) {
            this.processTextData();
        }
    }

    /* ---------- CONFIG ---------- */

    @api svgWidth = 800;
    @api svgHeight = 600;
    @api maxWords = 25; // Reduced to 25 for optimal spacing
    @api minWordLength = 3;
    @api textColumnName = 'text';
    @api showLegend;

    /** When 'dynamic', chart fills container and resizes with it. When 'fixed', uses Canvas Width/Height from properties. */
    _sizeMode = 'dynamic';

    @api
    get sizeMode() {
        return this._sizeMode;
    }
    set sizeMode(value) {
        this._sizeMode = value === 'fixed' ? 'fixed' : 'dynamic';
        if (this._sizeMode === 'fixed' && this._resizeObserver) {
            const container = this.template.querySelector('.visualization-container');
            if (container) this._resizeObserver.unobserve(container);
            this._resizeObserver = null;
        }
        if (this.d3Initialized && this.processedWords.length > 0) {
            setTimeout(() => this.renderVisualization(), 50);
        }
    }

    _chartType = CHART_TYPES.WORD_CLOUD;
    
    @api
    get chartType() {
        return this._chartType;
    }
    
    set chartType(value) {
        console.log('Setting chartType to:', value);
        this._chartType = value || CHART_TYPES.WORD_CLOUD;
        
        // If already initialized and have data, re-render
        if (this.d3Initialized && this.processedWords.length > 0) {
            setTimeout(() => {
                this.renderVisualization();
            }, 100);
        }
    }

    /* ---------- STATE ---------- */

    @track wordFrequency = new Map();
    @track processedWords = [];
    @track d3Initialized = false;
    @track isProcessing = false;
    @track currentOrientation = ORIENTATION.BOTH;
    @track error;
    @track _containerWidth = 0;
    @track _containerHeight = 0;

    _resizeObserver;
    _resizeDebounce;
    _lastRenderedWidth = 0;
    _lastRenderedHeight = 0;

    /* ---------- GETTERS ---------- */

    get hasData() {
        return Array.isArray(this.results) && this.results.length > 0;
    }

    get showVisualization() {
        return !this.isProcessing && this.processedWords.length > 0 && !this.error;
    }

    get totalRows() {
        return this.results?.length || 0;
    }

    get totalWordsProcessed() {
        return Array.from(this.wordFrequency.values()).reduce((sum, count) => sum + count, 0);
    }

    get uniqueWords() {
        return this.wordFrequency.size;
    }

    get displayedWords() {
        return this.processedWords.length;
    }

    get horizontalOnlyVariant() {
        return this.currentOrientation === ORIENTATION.HORIZONTAL ? 'brand' : 'neutral';
    }

    get verticalOnlyVariant() {
        return this.currentOrientation === ORIENTATION.VERTICAL ? 'brand' : 'neutral';
    }

    get bothOrientationVariant() {
        return this.currentOrientation === ORIENTATION.BOTH ? 'brand' : 'neutral';
    }

    get isWordCloud() {
        return [CHART_TYPES.WORD_CLOUD, CHART_TYPES.HORIZONTAL_WORD_CLOUD, 
                CHART_TYPES.VERTICAL_WORD_CLOUD, CHART_TYPES.MIXED_WORD_CLOUD].includes(this.chartType);
    }

    get isBarChart() {
        return [CHART_TYPES.HORIZONTAL_BAR, CHART_TYPES.VERTICAL_BAR].includes(this.chartType);
    }

    get showOrientationControls() {
        return this.chartType === CHART_TYPES.WORD_CLOUD;
    }

    /* ---------- THEME ---------- */

    @api theme = 'light'; // 'light' | 'night'

    get themeClass() {
        return this.theme === 'night' ? 'chart-wrapper night-theme' : 'chart-wrapper';
    }

    /* ---------- TITLES ---------- */

    @api chartTitleOverride; // optional custom title

    get chartTitle() {
        if (this.chartTitleOverride && this.chartTitleOverride.trim().length > 0) {
            return this.chartTitleOverride;
        }

        switch(this.chartType) {
            case CHART_TYPES.HORIZONTAL_BAR: return 'Horizontal Bar Chart';
            case CHART_TYPES.VERTICAL_BAR: return 'Vertical Bar Chart';
            case CHART_TYPES.HORIZONTAL_WORD_CLOUD: return 'Horizontal Word Cloud';
            case CHART_TYPES.VERTICAL_WORD_CLOUD: return 'Vertical Word Cloud';
            case CHART_TYPES.MIXED_WORD_CLOUD: return 'Mixed Word Cloud';
            default: return 'Word Cloud Analysis';
        }
    }

    get effectiveSvgWidth() {
        if (this.sizeMode === 'dynamic' && this._containerWidth > 0) {
            return this._containerWidth;
        }
        return Math.max(200, Number(this.svgWidth) || 800);
    }

    get effectiveSvgHeight() {
        if (this.sizeMode === 'dynamic' && this._containerHeight > 0) {
            return this._containerHeight;
        }
        return Math.max(200, Number(this.svgHeight) || 600);
    }

    /* ---------- LIFECYCLE ---------- */

    connectedCallback() {
        console.log('Component connected, chartType:', this.chartType);
    }

    disconnectedCallback() {
        if (this._resizeObserver) {
            const container = this.template.querySelector('.visualization-container');
            if (container) this._resizeObserver.unobserve(container);
            this._resizeObserver = null;
        }
        if (this._resizeDebounce) clearTimeout(this._resizeDebounce);
    }

    renderedCallback() {
        if (this.d3Initialized) {
            this._setupResizeObserver();
            // Fixed mode: re-render when canvas dimensions change from XML
            if (this.sizeMode === 'fixed' && this.processedWords.length > 0) {
                const w = this.effectiveSvgWidth;
                const h = this.effectiveSvgHeight;
                if (w !== this._lastRenderedWidth || h !== this._lastRenderedHeight) {
                    this._lastRenderedWidth = w;
                    this._lastRenderedHeight = h;
                    setTimeout(() => this.renderVisualization(), 0);
                }
            }
            return;
        }

        console.log('Loading D3 scripts...', 'chartType:', this.chartType);
        this.d3Initialized = true;

        // Always load both D3 scripts to prevent issues when chart type changes
        Promise.all([
            loadScript(this, D3 + '/d3/d3.min.js'),
            loadScript(this, D3 + '/d3/d3.layout.cloud.js')
        ])
            .then(() => {
                console.log('All D3 scripts loaded successfully');
                if (Array.isArray(this.results) && this.results.length > 0) {
                    this.processTextData();
                } else {
                    console.log('No results data yet');
                }
            })
            .catch(err => {
                this.error = 'Failed to load visualization library: ' + err.message;
                console.error('D3 load failed:', err);
            });
    }

    _setupResizeObserver() {
        if (this.sizeMode !== 'dynamic') return;
        const container = this.template.querySelector('.visualization-container');
        if (!container || this._resizeObserver) return;

        const measure = () => {
            const rect = container.getBoundingClientRect();
            const w = Math.floor(rect.width);
            const h = Math.floor(rect.height);
            if (w > 0 && h > 0 && (w !== this._containerWidth || h !== this._containerHeight)) {
                this._containerWidth = w;
                this._containerHeight = h;
                if (this.processedWords.length > 0) {
                    if (this._resizeDebounce) clearTimeout(this._resizeDebounce);
                    this._resizeDebounce = setTimeout(() => this.renderVisualization(), 150);
                }
            }
        };

        measure();
        this._resizeObserver = new ResizeObserver(measure);
        this._resizeObserver.observe(container);
    }

    /* ---------- EVENT HANDLERS ---------- */

    handleHorizontalOnly() {
        this.currentOrientation = ORIENTATION.HORIZONTAL;
        this.renderVisualization();
    }

    handleVerticalOnly() {
        this.currentOrientation = ORIENTATION.VERTICAL;
        this.renderVisualization();
    }

    handleBothOrientation() {
        this.currentOrientation = ORIENTATION.BOTH;
        this.renderVisualization();
    }

    /* ---------- TEXT PROCESSING ---------- */

    processTextData() {
        if (!Array.isArray(this.results) || this.results.length === 0) {
            console.warn('No results to process');
            return;
        }

        console.log('Processing text data from', this.results.length, 'rows');
        this.isProcessing = true;
        this.error = null;
        this.wordFrequency.clear();

        try {
            this.results.forEach((row) => {
                let textContent = '';

                if (typeof row === 'string') {
                    textContent = row;
                } else if (typeof row === 'object' && row !== null) {
                    const possibleKeys = [
                        this.textColumnName,
                        'text','content','description','message'
                    ];

                    for (const key of possibleKeys) {
                        if (row[key]) {
                            textContent = row[key];
                            break;
                        }
                    }

                    if (!textContent) {
                        textContent = Object.values(row).find(v => typeof v === 'string') || '';
                    }
                }

                if (textContent) {
                    const words = this.extractWords(textContent);
                    words.forEach(word => {
                        this.wordFrequency.set(
                            word,
                            (this.wordFrequency.get(word) || 0) + 1
                        );
                    });
                }
            });

            this.processedWords = Array.from(this.wordFrequency.entries())
                .map(([text, count]) => ({ text, count }))
                .sort((a, b) => b.count - a.count)
                .slice(0, this.maxWords);

            console.log('Processed', this.processedWords.length, 'unique words');
            this.isProcessing = false;
            
            // Use setTimeout to ensure DOM is ready
            setTimeout(() => {
                this.renderVisualization();
            }, 100);
        } catch (err) {
            console.error('Error processing text data:', err);
            this.error = 'Error processing text: ' + err.message;
            this.isProcessing = false;
        }
    }

    extractWords(text) {
        return text
            .toLowerCase()
            .replace(/[^\w\s]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim()
            .split(' ')
            .filter(word =>
                word.length >= this.minWordLength &&
                !STOP_WORDS.has(word)
            );
    }

    /* ---------- RENDER LOGIC ---------- */

    renderVisualization() {
        console.log('renderVisualization called, chartType:', this.chartType);
        console.log('processedWords length:', this.processedWords.length);
        
        if (!this.processedWords || this.processedWords.length === 0) {
            console.warn('No processed words available for rendering');
            return;
        }

        // In dynamic mode, ensure we have container dimensions before first paint
        if (this.sizeMode === 'dynamic') {
            const container = this.template.querySelector('.visualization-container');
            if (container) {
                const rect = container.getBoundingClientRect();
                const w = Math.floor(rect.width);
                const h = Math.floor(rect.height);
                if (w > 0 && h > 0) {
                    this._containerWidth = w;
                    this._containerHeight = h;
                }
            }
        }
        
        // Determine orientation based on chart type
        if (this.chartType === CHART_TYPES.HORIZONTAL_WORD_CLOUD) {
            this.currentOrientation = ORIENTATION.HORIZONTAL;
        } else if (this.chartType === CHART_TYPES.VERTICAL_WORD_CLOUD) {
            this.currentOrientation = ORIENTATION.VERTICAL;
        } else if (this.chartType === CHART_TYPES.MIXED_WORD_CLOUD) {
            this.currentOrientation = ORIENTATION.BOTH;
        }

        try {
            if (this.isBarChart) {
                console.log('Rendering bar chart');
                this.renderBarChart();
            } else {
                console.log('Rendering word cloud');
                this.renderWordCloud();
            }
        } catch (err) {
            console.error('Error rendering visualization:', err);
            this.error = 'Error rendering chart: ' + err.message;
        }
    }

    /* ---------- WORD CLOUD RENDER ---------- */

    getRotation() {
        switch(this.currentOrientation) {
            case ORIENTATION.HORIZONTAL:
                return () => 0;
            case ORIENTATION.VERTICAL:
                return () => 90;
            case ORIENTATION.BOTH:
            default:
                return () => Math.random() > 0.5 ? 0 : 90;
        }
    }

    renderWordCloud() {
        if (!this.processedWords.length) {
            console.warn('No processed words to render');
            return;
        }

        const svg = this.template.querySelector('svg.word-cloud-svg');
        if (!svg) {
            console.error('SVG element not found');
            return;
        }

        // Check if d3.layout.cloud is available
        // eslint-disable-next-line no-undef
        if (typeof d3 === 'undefined') {
            console.error('D3 not loaded');
            this.error = 'Visualization library not loaded. Please refresh the page.';
            return;
        }

        // eslint-disable-next-line no-undef
        if (typeof d3.layout === 'undefined' || typeof d3.layout.cloud === 'undefined') {
            console.error('D3 word cloud layout not loaded');
            this.error = 'Word cloud library not loaded. Please refresh the page.';
            return;
        }

        console.log('Rendering word cloud with', this.processedWords.length, 'words');

        // eslint-disable-next-line no-undef
        d3.select(svg).selectAll('*').remove();

        svg.setAttribute('width', this.effectiveSvgWidth);
        svg.setAttribute('height', this.effectiveSvgHeight);

        const isHorizontalCloud = this.chartType === CHART_TYPES.HORIZONTAL_WORD_CLOUD || this.currentOrientation === ORIENTATION.HORIZONTAL;

        const max = Math.max(...this.processedWords.map(w => w.count));
        const min = Math.min(...this.processedWords.map(w => w.count));

        // Use slightly smaller fonts and fewer words for strictly horizontal layouts
        const baseWords = isHorizontalCloud ? this.processedWords.slice(0, Math.min(18, this.processedWords.length)) : this.processedWords;

        const words = baseWords.map((w, i) => {
            const baseMin = isHorizontalCloud ? 10 : 12;
            const baseRange = isHorizontalCloud ? 20 : 36; // max ~30px vs ~48px

            return {
                text: w.text,
                count: w.count,
                size: baseMin + ((w.count - min) / (max - min || 1)) * baseRange,
                color: COLOR_PALETTE[i % COLOR_PALETTE.length]
            };
        });

        // eslint-disable-next-line no-undef
        // Reserve some vertical space for legend at the top to avoid overlap
        const legendReserve = this.showLegend ? 120 : 60;

        const layout = d3.layout.cloud()
            .size([
                this.effectiveSvgWidth - 60,
                Math.max(120, this.effectiveSvgHeight - legendReserve)
            ])
            .words(words)
            .padding(isHorizontalCloud ? 10 : 28)
            .rotate(this.getRotation())
            .fontSize(d => d.size)
            .spiral('archimedean') // FIXED: Changed back to archimedean for better packing
            .timeInterval(100) // FIXED: Increased time for better collision detection
            .random(() => 0.5) // Consistent random
            .on('end', words => {
                console.log('Word cloud layout complete, positioned', words.length, 'words');
                this.drawWordCloud(words, svg);
            });

        // FIXED: Add error handling for layout
        try {
            layout.start();
        } catch (err) {
            console.error('Word cloud layout error:', err);
            this.error = 'Error creating word cloud layout. Try reducing the number of words.';
        }
    }

    drawWordCloud(words, svg) {
        console.log('Drawing', words.length, 'words that were successfully positioned');
        
        // eslint-disable-next-line no-undef
        const g = d3.select(svg)
            .append('g')
            .attr(
                'transform',
                `translate(${this.effectiveSvgWidth / 2},${this.effectiveSvgHeight / 2})`
            );

        const tooltip = this.template.querySelector('[data-tooltip]');

        // Add legend if enabled
        if (this.showLegend) {
            this.addLegend(svg);
        }

        // Filter out words that weren't positioned (x and y are undefined)
        const positionedWords = words.filter(d => d.x !== undefined && d.y !== undefined);
        console.log('Successfully positioned:', positionedWords.length, 'out of', words.length);

        // FIXED: Add warning if many words couldn't be positioned
        if (positionedWords.length < words.length * 0.7) {
            console.warn(`Only ${positionedWords.length} of ${words.length} words could be positioned. Consider reducing maxWords or increasing canvas size.`);
        }

        g.selectAll('text')
            .data(positionedWords)
            .enter()
            .append('text')
            .style('font-size', d => `${d.size}px`)
            .style('font-family', 'Arial, sans-serif')
            .style('font-weight', '600')
            .style('fill', d => d.color)
            .style('cursor', 'pointer')
            .style('transition', 'all 0.3s ease')
            .style('user-select', 'none')
            .attr('text-anchor', 'middle')
            .attr('transform', d => `translate(${d.x},${d.y})rotate(${d.rotate || 0})`)
            .text(d => d.text)
            .on('mouseover', function(event, d) {
                // eslint-disable-next-line no-undef
                d3.select(this)
                    .style('fill', '#ff6b6b')
                    .style('font-size', `${d.size * 1.1}px`); // FIXED: Reduced hover scale from 1.15 to 1.1
                
                if (tooltip) {
                    tooltip.style.display = 'block';
                    tooltip.innerHTML = `
                        <div class="tooltip-content">
                            <strong>${d.text}</strong>
                            <div class="tooltip-count">Count: ${d.count}</div>
                        </div>
                    `;
                }
            })
            .on('mousemove', function(event) {
                if (tooltip) {
                    tooltip.style.left = `${event.pageX + 15}px`;
                    tooltip.style.top = `${event.pageY + 15}px`;
                }
            })
            .on('mouseout', function(event, d) {
                // eslint-disable-next-line no-undef
                d3.select(this)
                    .style('fill', d.color)
                    .style('font-size', `${d.size}px`);
                
                if (tooltip) {
                    tooltip.style.display = 'none';
                }
            });
    }

    /* ---------- BAR CHART RENDER ---------- */

    renderBarChart() {
        if (!this.processedWords.length) {
            console.warn('No processed words to render bar chart');
            return;
        }

        const svg = this.template.querySelector('svg.word-cloud-svg');
        if (!svg) {
            console.error('SVG element not found for bar chart');
            return;
        }

        // Check if d3 is available
        // eslint-disable-next-line no-undef
        if (typeof d3 === 'undefined') {
            console.error('D3 not loaded');
            this.error = 'D3 library not loaded. Please refresh the page.';
            return;
        }

        console.log('Rendering bar chart with', this.processedWords.length, 'words');

        // eslint-disable-next-line no-undef
        d3.select(svg).selectAll('*').remove();

        const isVertical = this.chartType === CHART_TYPES.VERTICAL_BAR;
        
        const margin = { top: 40, right: 40, bottom: isVertical ? 100 : 120, left: isVertical ? 120 : 100 };
        const width = this.effectiveSvgWidth - margin.left - margin.right;
        const height = this.effectiveSvgHeight - margin.top - margin.bottom;

        svg.setAttribute('width', this.effectiveSvgWidth);
        svg.setAttribute('height', this.effectiveSvgHeight);

        // eslint-disable-next-line no-undef
        const g = d3.select(svg)
            .append('g')
            .attr('transform', `translate(${margin.left},${margin.top})`);

        const data = this.processedWords.slice(0, 20);
        console.log('Bar chart data:', data.length, 'items');

        if (isVertical) {
            this.renderVerticalBarChart(g, data, width, height);
        } else {
            this.renderHorizontalBarChart(g, data, width, height);
        }

        if (this.showLegend) {
            this.addLegend(svg);
        }
    }

    renderVerticalBarChart(g, data, width, height) {
        // eslint-disable-next-line no-undef
        const x = d3.scaleBand()
            .domain(data.map(d => d.text))
            .range([0, width])
            .padding(0.2);

        // eslint-disable-next-line no-undef
        const y = d3.scaleLinear()
            // eslint-disable-next-line no-undef
            .domain([0, d3.max(data, d => d.count)])
            .range([height, 0]);

        g.append('g')
            .attr('transform', `translate(0,${height})`)
            // eslint-disable-next-line no-undef
            .call(d3.axisBottom(x))
            .selectAll('text')
            .attr('transform', 'rotate(-45)')
            .style('text-anchor', 'end')
            .style('font-size', '11px')
            .style('fill', '#374151');

        g.append('g')
            // eslint-disable-next-line no-undef
            .call(d3.axisLeft(y).ticks(5))
            .selectAll('text')
            .style('font-size', '11px')
            .style('fill', '#374151');

        const tooltip = this.template.querySelector('[data-tooltip]');

        g.selectAll('.bar')
            .data(data)
            .enter()
            .append('rect')
            .attr('class', 'bar')
            .attr('x', d => x(d.text))
            .attr('y', d => y(d.count))
            .attr('width', x.bandwidth())
            .attr('height', d => height - y(d.count))
            .attr('fill', (d, i) => COLOR_PALETTE[i % COLOR_PALETTE.length])
            .attr('rx', 4)
            .style('cursor', 'pointer')
            .on('mouseover', function(event, d) {
                // eslint-disable-next-line no-undef
                d3.select(this).attr('opacity', 0.7);
                
                if (tooltip) {
                    tooltip.style.display = 'block';
                    tooltip.innerHTML = `
                        <div class="tooltip-content">
                            <strong>${d.text}</strong>
                            <div class="tooltip-count">Count: ${d.count}</div>
                        </div>
                    `;
                }
            })
            .on('mousemove', function(event) {
                if (tooltip) {
                    tooltip.style.left = `${event.pageX + 15}px`;
                    tooltip.style.top = `${event.pageY + 15}px`;
                }
            })
            .on('mouseout', function() {
                // eslint-disable-next-line no-undef
                d3.select(this).attr('opacity', 1);
                if (tooltip) {
                    tooltip.style.display = 'none';
                }
            });

        g.selectAll('.label')
            .data(data)
            .enter()
            .append('text')
            .attr('class', 'label')
            .attr('x', d => x(d.text) + x.bandwidth() / 2)
            .attr('y', d => y(d.count) - 5)
            .attr('text-anchor', 'middle')
            .style('font-size', '10px')
            .style('font-weight', '600')
            .style('fill', '#374151')
            .text(d => d.count);
    }

    renderHorizontalBarChart(g, data, width, height) {
        // eslint-disable-next-line no-undef
        const y = d3.scaleBand()
            .domain(data.map(d => d.text))
            .range([0, height])
            .padding(0.2);

        // eslint-disable-next-line no-undef
        const x = d3.scaleLinear()
            // eslint-disable-next-line no-undef
            .domain([0, d3.max(data, d => d.count)])
            .range([0, width]);

        g.append('g')
            // eslint-disable-next-line no-undef
            .call(d3.axisLeft(y))
            .selectAll('text')
            .style('font-size', '11px')
            .style('fill', '#374151');

        g.append('g')
            .attr('transform', `translate(0,${height})`)
            // eslint-disable-next-line no-undef
            .call(d3.axisBottom(x).ticks(5))
            .selectAll('text')
            .style('font-size', '11px')
            .style('fill', '#374151');

        const tooltip = this.template.querySelector('[data-tooltip]');

        g.selectAll('.bar')
            .data(data)
            .enter()
            .append('rect')
            .attr('class', 'bar')
            .attr('y', d => y(d.text))
            .attr('x', 0)
            .attr('height', y.bandwidth())
            .attr('width', d => x(d.count))
            .attr('fill', (d, i) => COLOR_PALETTE[i % COLOR_PALETTE.length])
            .attr('rx', 4)
            .style('cursor', 'pointer')
            .on('mouseover', function(event, d) {
                // eslint-disable-next-line no-undef
                d3.select(this).attr('opacity', 0.7);
                
                if (tooltip) {
                    tooltip.style.display = 'block';
                    tooltip.innerHTML = `
                        <div class="tooltip-content">
                            <strong>${d.text}</strong>
                            <div class="tooltip-count">Count: ${d.count}</div>
                        </div>
                    `;
                }
            })
            .on('mousemove', function(event) {
                if (tooltip) {
                    tooltip.style.left = `${event.pageX + 15}px`;
                    tooltip.style.top = `${event.pageY + 15}px`;
                }
            })
            .on('mouseout', function() {
                // eslint-disable-next-line no-undef
                d3.select(this).attr('opacity', 1);
                if (tooltip) {
                    tooltip.style.display = 'none';
                }
            });

        g.selectAll('.label')
            .data(data)
            .enter()
            .append('text')
            .attr('class', 'label')
            .attr('y', d => y(d.text) + y.bandwidth() / 2)
            .attr('x', d => x(d.count) + 5)
            .attr('dy', '0.35em')
            .style('font-size', '10px')
            .style('font-weight', '600')
            .style('fill', '#374151')
            .text(d => d.count);
    }

    /* ---------- LEGEND ---------- */

    addLegend(svg) {
        // eslint-disable-next-line no-undef
        const legend = d3.select(svg)
            .append('g')
            .attr('class', 'legend')
            .attr('transform', 'translate(20, 20)');

        legend.append('text')
            .attr('x', 0)
            .attr('y', 0)
            .style('font-size', '12px')
            .style('font-weight', '600')
            .style('fill', '#6b7280')
            .text('Statistics');

        const stats = [
            { label: 'Total Words', value: this.totalWordsProcessed, color: '#667eea' },
            { label: 'Unique Words', value: this.uniqueWords, color: '#764ba2' },
            { label: 'Displayed', value: this.displayedWords, color: '#f093fb' }
        ];

        stats.forEach((stat, i) => {
            const group = legend.append('g')
                .attr('transform', `translate(0, ${20 + (i * 24)})`);

            group.append('rect')
                .attr('width', 12)
                .attr('height', 12)
                .attr('rx', 2)
                .attr('fill', stat.color);

            group.append('text')
                .attr('x', 18)
                .attr('y', 10)
                .style('font-size', '11px')
                .style('font-weight', '400')
                .style('fill', '#374151')
                .text(`${stat.label}: ${stat.value}`);
        });

        const bbox = legend.node().getBBox();
        legend.insert('rect', ':first-child')
            .attr('x', bbox.x - 8)
            .attr('y', bbox.y - 6)
            .attr('width', bbox.width + 16)
            .attr('height', bbox.height + 12)
            .attr('rx', 6)
            .attr('fill', 'white')
            .attr('stroke', '#e5e7eb')
            .attr('stroke-width', 1)
            .style('filter', 'drop-shadow(0 2px 4px rgba(0,0,0,0.08))');
    }
}