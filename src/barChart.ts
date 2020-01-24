/*
 *  Power BI Visual CLI
 *
 *  Copyright (c) Microsoft Corporation
 *  All rights reserved.
 *  MIT License
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the ""Software""), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in
 *  all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
 *  THE SOFTWARE.
 */

module powerbi.extensibility.visual {

    import TextProperties = powerbi.extensibility.utils.formatting.TextProperties;
    import textMeasurementService = powerbi.extensibility.utils.formatting.textMeasurementService;
    import IColorPalette = powerbi.extensibility.IColorPalette;
    import valueFormatter = powerbi.extensibility.utils.formatting.valueFormatter;
    import tooltip = powerbi.extensibility.utils.tooltip;
    import ITooltipServiceWrapper = powerbi.extensibility.utils.tooltip.ITooltipServiceWrapper;
    import createTooltipServiceWrapper = powerbi.extensibility.utils.tooltip.createTooltipServiceWrapper;
    export interface TooltipEventArgs<TData> {
        data: TData;
        coordinates: number[];
        elementCoordinates: number[];
        context: HTMLElement;
        isTouchEvent: boolean;
    }
    export module DataViewObjects {
        // Gets the value of the given object/property pair.
        // tslint:disable-next-line:no-shadowed-variable
        export function getValue<T>(objects: DataViewObjects, propertyId: DataViewObjectPropertyIdentifier, defaultValue?: T): T {

            if (!objects) {

                return defaultValue;
            }

            let objectOrMap: DataViewObject;
            objectOrMap = objects[propertyId.objectName];

            let object: DataViewObject;
            object = <DataViewObject>objectOrMap;

            return DataViewObject.getValue(object, propertyId.propertyName, defaultValue);
        }
        // Gets the solid color from a fill property. 
        export function getFillColor(objects: DataViewObjects,
                                     propertyId: DataViewObjectPropertyIdentifier, defaultColor?: string): string {
            let value: Fill;
            value = getValue(objects, propertyId);
            if (!value || !value.solid) {
                return defaultColor;
            }

            return value.solid.color;
        }
    }
    export module DataViewObject {
        // tslint:disable-next-line:no-shadowed-variable
        export function getValue<T>(object: DataViewObject, propertyName: string, defaultValue?: T): T {

            if (!object) {
                return defaultValue;
            }

            let propertyValue: T;
            propertyValue = <T>object[propertyName];
            if (propertyValue === undefined) {
                return defaultValue;
            }

            return propertyValue;
        }
    }

    interface IBarChartViewModel {
        dataPoints: IBarChartDataPoint[];
        dataMax: number;
        name: string;
        dataMin: number;
    }
    interface IBarChartDataPoint {
        value: number;
        category: string;
        format: string;
        color: string;
        selectionId: powerbi.visuals.ISelectionId;
    }
    export interface IMeasureTitle {
        fontSize: number;
        color: string;
    }
    export interface ILabelSettings {
        fontSize: number;
        color: string;
        displayUnits: number;
        strokeWidth: number;
    }
    export interface IAnimationSettings {
        show: boolean;
        duration: number;
    }

    // tslint:disable-next-line:typedef
    let props;
    props = {
        measureTitle: {
            color: <DataViewObjectPropertyIdentifier>{ objectName: 'measureTitle', propertyName: 'color' },
            fontSize: <DataViewObjectPropertyIdentifier>{ objectName: 'measureTitle', propertyName: 'fontSize' }
        },
        labelSettings: {
            color: <DataViewObjectPropertyIdentifier>{ objectName: 'labelSettings', propertyName: 'color' },
            fontSize: <DataViewObjectPropertyIdentifier>{ objectName: 'labelSettings', propertyName: 'fontSize' },
            displayUnits: <DataViewObjectPropertyIdentifier>{ objectName: 'labelSettings', propertyName: 'displayUnits' },
            strokeWidth: <DataViewObjectPropertyIdentifier>{ objectName: 'labelSettings', propertyName: 'strokeWidth' }
        },
        animationSettings: {
            show: <DataViewObjectPropertyIdentifier>{ objectName: 'animationSettings', propertyName: 'show' },
            duration: <DataViewObjectPropertyIdentifier>{ objectName: 'animationSettings', propertyName: 'duration' }
        }
    };

    function getValue<T>(objects: DataViewObjects, objectName: string, propertyName: string, defaultValue: T): T {
        if (objects) {
            let object: DataViewObject;
            object = objects[objectName];
            if (object) {
                let property: T;
                property = <T>object[propertyName];
                if (property !== undefined) {
                    return property;
                }
            }
        }

        return defaultValue;
    }

    function getCategoricalObjectValue<T>(category: DataViewCategoryColumn,
                                          index: number, objectName: string, propertyName: string, defaultValue: T): T {
        let categoryObjects: DataViewObjects[];
        categoryObjects = category.objects;

        if (categoryObjects) {
            let categoryObject: DataViewObject;
            categoryObject = categoryObjects[index];
            if (categoryObject) {
                let object: DataViewPropertyValue;
                object = categoryObject[objectName];
                if (object) {
                    let property: T;
                    property = object[propertyName];
                    if (property !== undefined) {
                        return property;
                    }
                }
            }
        }

        return defaultValue;
    }

    function visualTransform(options: VisualUpdateOptions, host: IVisualHost, measureIndex: number): IBarChartViewModel {
        let dataViews: DataView[];
        dataViews = options.dataViews;

        let viewModel: IBarChartViewModel;
        viewModel = {
            dataPoints: [],
            dataMax: 0,
            name: '',
            dataMin: 0
        };

        if (!dataViews || !dataViews[0] || !dataViews[0].categorical ||
            !dataViews[0].categorical.categories || !dataViews[0].categorical.categories[0].source || !dataViews[0].categorical.values) {
            return viewModel;
        }

        if (measureIndex > options.dataViews[0].categorical.values.length - 1) {
            measureIndex = 0;
        }
        let categorical: DataViewCategorical;
        categorical = dataViews[0].categorical;
        // tslint:disable-next-line:no-any
        let category: any;
        category = categorical.categories[0];
        // tslint:disable-next-line:no-any
        let dataValue: any;
        dataValue = categorical.values[measureIndex];
        let dataMax: number;
        // tslint:disable-next-line:no-any
        let mName: any;
        mName = categorical.values[measureIndex].source.displayName;

        let barChartDataPoints: IBarChartDataPoint[];
        barChartDataPoints = [];
        //let dataMax: number;
        let dataMin: number;
        let colorPalette: powerbi.extensibility.IColorPalette;
        colorPalette = host.colorPalette;
        let objects: DataViewObjects;
        objects = dataViews[0].metadata.objects;
        let len: number;
        len = Math.max(category.values.length, dataValue.values.length);

        for (let i: number = 0; i < len; i++) {
            let defaultColor: Fill;
            defaultColor = {
                solid: {
                    color: colorPalette.getColor(<string>category.values[i]).value
                }
            };

            barChartDataPoints.push({
                category: <string>category.values[i],
                value: <number>dataValue.values[i],
                format: dataValue.source.format,
                color: getCategoricalObjectValue<Fill>(category, i, 'colorSelector', 'fill', defaultColor).solid.color,
                selectionId: host.createSelectionIdBuilder()
                    .withCategory(category, i)
                    .createSelectionId()

            });
        }
        dataMax = <number>dataValue.maxLocal;

        dataMin = dataValue.minLocal;

        return {
            dataPoints: barChartDataPoints,
            dataMax: dataMax,
            name: mName,
            dataMin: dataMin
        };
    }

    export class BarChart implements IVisual {
        private target: HTMLElement;
        private updateCount: number;
        private svg: d3.Selection<SVGElement>;
        private div1: d3.Selection<SVGElement>;
        private div2: d3.Selection<SVGElement>;
        private host: IVisualHost;
        private events: IVisualEventService ;
        private horizBarChartContainer: d3.Selection<SVGElement>;
        private horizBarContainer: d3.Selection<SVGElement>;
        private measureTitle: d3.Selection<SVGElement>;
        private horizBars: d3.Selection<SVGElement>;
        private selectionManager: ISelectionManager;
        private barDataPoints: IBarChartDataPoint[];
        private yAxis: d3.Selection<SVGElement>;
        private tooltipServiceWrapper: ITooltipServiceWrapper;
        private yAxisMeasures: d3.Selection<SVGElement>;
        private rotationId: number;
        private frameId: number;
        private measureUpdateCounter: number;
        private viewModel: IBarChartViewModel;
        private width: number;
        private height: number;
        private margin: number;
        // tslint:disable-next-line:typedef
        private xScale;
        private rotationCount: number;
        // tslint:disable-next-line:typedef
        private horizBarChart;
        private options: VisualUpdateOptions;
        private measureCount: number;
        private dataviews: DataView;
        private rootElement: d3.Selection<SVGElement>;
        private maxData: number;
        private minData: number;

        // tslint:disable-next-line:typedef
        public static statConfig = {
            xScalePadding: 0.1,
            solidOpacity: 1,
            transparentOpacity: 0.5,
            margins: {
                top: 0,
                right: 20,
                bottom: 25,
                left: 20
            }
        };

        constructor(options: VisualConstructorOptions) {
            this.events=options.host.eventService;
            this.measureUpdateCounter = 0;
            this.host = options.host;
            this.selectionManager = options.host.createSelectionManager();
            this.tooltipServiceWrapper = createTooltipServiceWrapper(this.host.tooltipService, options.element);
            this.rootElement = d3.select(options.element);
            this.div1 = this.rootElement.append('div')
                .classed('rootDiv', true).style('xoverflow', 'visible');
            this.div2 = this.div1.append('div')
                .classed('baseDiv', true);

            this.svg = this.div2
                .append('svg')
                .classed('horizBarChart', true);

            this.measureTitle = this.svg.append('text')
                .classed('measureTitle', true);

            this.horizBarContainer = this.svg.append('g')
                .classed('horizBarContainer', true);

            this.yAxis = this.svg.append('g')
                .classed('yAxis', true);

            this.yAxisMeasures = this.svg.append('g')
                .classed('yAxisMeasures', true);

        }

        public update(options: VisualUpdateOptions): void {
            this.options = options;
            this.events.renderingStarted(options);
            this.dataviews = options.dataViews[0];

            clearInterval(this.frameId);
            clearInterval(this.rotationId);
            this.horizBarContainer.selectAll('*').remove();
            this.yAxis.selectAll('*').remove();
            this.yAxisMeasures.selectAll('*').remove();

            this.viewModel = visualTransform(options, this.host, this.measureUpdateCounter);
            this.barDataPoints = this.viewModel.dataPoints;

            let measureTitle: IMeasureTitle;
            measureTitle = this.getMeasureTitle(this.dataviews);
            let animationSettings: IAnimationSettings;
            animationSettings = this.getAnimationSettings(this.dataviews);
            let labelSettings: ILabelSettings;
            labelSettings = this.getLabelSettings(this.dataviews);

            this.svg.attr({
                width: options.viewport.width,
                height: options.viewport.height
            });

            this.horizBarChart = $('.horizBarChart');
            this.horizBarChart.css('transform', `rotateX(0deg)`);
            this.width = options.viewport.width;
            this.height = options.viewport.height;

            this.yAxis.style({
                'font-size': `${labelSettings.fontSize}px`, fill: labelSettings.color
            });
            this.yAxisMeasures.style({
                'font-size': `${labelSettings.fontSize}px`, fill: labelSettings.color
            });

            let titleProperties: TextProperties;
            titleProperties = {
                text: this.viewModel.name,
                fontFamily: 'Segoe UI,wf_segoe-ui_normal,helvetica,arial,sans-serif',
                fontSize: `${measureTitle.fontSize}px`
            };
            // measure title
            this.svg.select('.measureTitle')
                .attr('transform', `translate(20,${(measureTitle.fontSize)})`)
                .attr('font-size', `${measureTitle.fontSize}px`)
                .attr('fill', measureTitle.color);
            let textHeight: number;
            textHeight = textMeasurementService.measureSvgTextHeight(titleProperties);
            this.rootElement.select('.rootDiv').style('height', `${options.viewport.height}px`);
            this.rootElement.select('.baseDiv').style('width', '100%');
            this.margin = 15 / 100;
            this.xScale = d3.scale.ordinal()
                // tslint:disable-next-line:typedef
                .domain(this.viewModel.dataPoints.map(d => d.category))
                .rangeBands([BarChart.statConfig.margins.top + textHeight, this.height], 0.2, 0.3);

            let barHeight: number;
            barHeight = this.xScale.rangeBand();

            if (barHeight < 20) {
                this.height = options.viewport.height + (this.viewModel.dataPoints.length * (20 - barHeight));
                this.width = options.viewport.width - 20;
                this.xScale.rangeBands([BarChart.statConfig.margins.top + textHeight, this.height], 0.2, 0.3);
                this.div1.select('.baseDiv').style('height', `${this.height}px`);
                this.div1.select('.horizBarChart').style('height', `${this.height}px`);
            } else {
                this.height = options.viewport.height;
                this.width = options.viewport.width;
                this.xScale.rangeBands([BarChart.statConfig.margins.top + textHeight, this.height], 0.2, 0.3);
                this.div1.select('.baseDiv').style('height', `${this.height}px`);
                this.div1.select('.horizBarChart').style('height', `${this.height}px`);
            }
             let yAxis: d3.svg.Axis;
            yAxis = d3.svg.axis()
                .scale(this.xScale)
                .orient('left');
             this.yAxis.attr('transform', `translate(${this.margin * this.width},0)`)
                .call(yAxis);
            this.measureCount = options.dataViews[0].categorical.values.length;
            this.renderVisual();
             if (this.measureCount > 1) {
                if (animationSettings.show) {
                    clearInterval(this.rotationId);
                    this.rotationId = setInterval(() => this.rotation(), animationSettings.duration * 1000);
                }
                // Click functionality
                $('.horizBarChart').on('click', () => {
                    clearInterval(this.rotationId);
                    this.rotation();
                    if (animationSettings.show) {
                        clearInterval(this.rotationId);
                        this.rotationId = setInterval(() => this.rotation(), animationSettings.duration * 1000);
                    }
                });
            }this.events.renderingFinished(options);
        }

        public rotation(): void {
            this.rotationCount = 1;
            clearInterval(this.frameId);
            this.frameId = setInterval(() => this.frame(), 5);
        }

        public frame(): void {
            if (this.rotationCount === 90) {
                this.measureUpdateCounter++;
                if (this.measureUpdateCounter >= this.measureCount) {
                    this.measureUpdateCounter = 0;
                }

                this.rotationCount = -90;
                this.viewModel = visualTransform(this.options, this.host, this.measureUpdateCounter);
                this.renderVisual();
            } else if (this.rotationCount === 0) {
                clearInterval(this.frameId);
            } else {
                this.rotationCount++;
                this.horizBarChart.css('transform', `rotateX(${-this.rotationCount}deg)`);
            }
        }

        public renderVisual(): void {
            let THIS: this;
            THIS = this;
            let width: number;
            width = this.width;
            // tslint:disable-next-line
            const availableWidth: any = width / 9;
            let measureTitle: IMeasureTitle;
            measureTitle = this.getMeasureTitle(this.dataviews);
            let labelSettings: ILabelSettings;
            labelSettings = this.getLabelSettings(this.dataviews);
            let yScale: d3.scale.Linear<number, number>;
            if (this.viewModel.dataMin < 0) {
                yScale = d3.scale.linear()
                    .domain([0, Math.abs(this.viewModel.dataMax) + Math.abs(this.viewModel.dataMin)])
                    .range([this.width, (this.margin * this.width * 2)]);
            } else {
                yScale = d3.scale.linear()
                    .domain([0, Math.abs(this.viewModel.dataMax)])
                    .range([this.width, (this.margin * this.width * 2)]);

            }

            let titleProperties: TextProperties;
            titleProperties = {
                text: this.viewModel.name,
                fontFamily: 'Segoe UI,wf_segoe-ui_normal,helvetica,arial,sans-serif',
                fontSize: `${measureTitle.fontSize}px`
            };

            $('.measureTitle').text(textMeasurementService.getTailoredTextOrDefault(titleProperties, this.width - 30));

            let add: number;
            if ((this.viewModel.dataMax > 0 && this.viewModel.dataMin < 0)
                || ((this.viewModel.dataMax < 0 && this.viewModel.dataMin < 0))) {

                add = (this.width - yScale(Math.abs(this.viewModel.dataMin)));
            } else {
                add = 0;
            }
            let bars: d3.selection.Update<IBarChartDataPoint>;
            bars = this.horizBarContainer.selectAll('.bar').data(this.viewModel.dataPoints);
            bars.enter()
                .append('rect')
                .classed('bar', true);
            // bars
            bars.attr({
                // tslint:disable-next-line:typedef
                width: d => (this.width - yScale(parseFloat(`${d.value}`))) < 0 ? (this.width - yScale(d.value * -1))
                    : (this.width - yScale(d.value)),
                height: this.xScale.rangeBand(),
                // tslint:disable-next-line:typedef
                y: d => this.xScale(d.category),
                // tslint:disable-next-line:typedef
                x: d => (this.width - yScale(parseFloat(`${d.value}`))) < 0 ?
                    (d.value === this.viewModel.dataMin ? (this.width * this.margin) : (this.width * this.margin) + Math.abs(add)
                        - ((this.width - yScale(parseFloat(`${d.value}`))) < 0 ? (this.width - yScale(d.value * -1))
                            : (this.width - yScale(d.value)))) :
                    (this.width * this.margin) + Math.abs(add),
                // tslint:disable-next-line:typedef
                fill: d => d.color,
                'fill-opacity': BarChart.statConfig.solidOpacity
            });

            let barHeight: number;
            barHeight = +bars.attr('height');
            if (this.width > 90) {
                this.yAxisMeasures.selectAll('text').remove();
                let measureValue: d3.selection.Update<IBarChartDataPoint>;
                measureValue = this.yAxisMeasures.selectAll('text').data(this.viewModel.dataPoints);
                let measureLabel: d3.Selection<IBarChartDataPoint>;

                const format: string = THIS.viewModel.dataPoints[0].format;
                // tslint:disable-next-line:typedef
                let formatter;
                let tempMeasureData: String;
                tempMeasureData = Math.round(THIS.viewModel.dataMax).toString();
                let displayVal: number = 0;
                if (labelSettings.displayUnits === 0) {
                    let valLen: number;
                    valLen = tempMeasureData.length;
                    if (valLen > 9) {
                        displayVal = 1e9;
                    } else if (valLen <= 9 && valLen > 6) {
                        displayVal = 1e6;
                    } else if (valLen <= 6 && valLen >= 4) {
                        displayVal = 1e3;
                    } else {
                        displayVal = 10;
                    }
                }
                if (format && format.indexOf('%') !== -1) {
                    formatter = valueFormatter.create({
                        format: format,
                        value: labelSettings.displayUnits === 0 ? 0 : labelSettings.displayUnits,
                        precision: labelSettings.strokeWidth
                    });
                } else {
                    formatter = valueFormatter.create({
                        format: format,
                        value: labelSettings.displayUnits === 0 ? displayVal : labelSettings.displayUnits,
                        precision: labelSettings.strokeWidth
                    });
                }
                measureLabel = measureValue.enter()
                    .append('text')
                    .classed('measureValue', true);
                // measure value
                measureLabel.attr({
                    dy: '0.40em',
                    // tslint:disable-next-line:typedef
                    y: d => this.xScale(d.category) + (barHeight * 0.5),
                    // tslint:disable-next-line:typedef
                    x: d => this.width - (this.margin * this.width) + 10
                })
                .text((d: IBarChartDataPoint): string => {
                    return THIS.applyEllipsis(d.value, formatter, labelSettings, availableWidth, measureValue);

                })
                .append('title').text((d: IBarChartDataPoint): string => {
                    return formatter.format(d.value);
                });
            }

            // Changing the text to ellipsis if the width of the window is small
            for (let i: number = 0; i < this.viewModel.dataPoints.length; i++) {
                let newDataLabel: string;
                let ticktxt: any;
                ticktxt = $('.tick text');
                let tick: any;
                tick=ticktxt[i];
                newDataLabel = THIS.applyEllipsis(this.viewModel.dataPoints[i].category, null, labelSettings, availableWidth, null);
                if (ticktxt && tick) {
                    tick.textContent = newDataLabel;
                    d3.select(tick).append('title').text(this.viewModel.dataPoints[i].category);
                    d3.select(tick).attr('line-height', '10px');
                }
            }

            this.tooltipServiceWrapper.addTooltip(this.horizBarContainer.selectAll('.bar'),
                                                  (tooltipEvent: TooltipEventArgs<IBarChartDataPoint>) =>
                                                   BarChart.getTooltipData(tooltipEvent.data),
                                                  (tooltipEvent: TooltipEventArgs<IBarChartDataPoint>) => tooltipEvent.data.selectionId);

            let selectionManager: ISelectionManager;
            selectionManager = this.selectionManager;
            bars.exit()
                .remove();
        }

        // tslint:disable-next-line:no-any
        public applyEllipsis(d: any, formatter: any, labelSettings: ILabelSettings, width: any, measureValue: any): string {
            let measureProperties: TextProperties;
            measureProperties = {
                text: formatter === null ? d : formatter.format(d),
                fontFamily: 'sans-serif',
                fontSize: `${labelSettings.fontSize}px`
            };

            return textMeasurementService.getTailoredTextOrDefault(measureProperties, width);
        }

        private getDefaultMeasureTitle(): IMeasureTitle {
            return {
                color: '#666666',
                fontSize: 20
            };
        }

        private getDefaultLabelSettings(): ILabelSettings {
            return {
                color: '#000',
                fontSize: 12,
                displayUnits: 0,
                strokeWidth: 0
            };
        }

        public getDefaultAnimationSettings(): IAnimationSettings {
            return {
                show: false,
                duration: 6
            };
        }

        private getMeasureTitle(dataView: DataView): IMeasureTitle {
            let objects: DataViewObjects = null;
            let title: IMeasureTitle;
            title = this.getDefaultMeasureTitle();
            if (!dataView.metadata || !dataView.metadata.objects) {
                return title;
            }
            objects = dataView.metadata.objects;
            // tslint:disable-next-line:typedef
            const currentmeasurelabelprop = props;
            title.color = DataViewObjects.getFillColor(objects, currentmeasurelabelprop.measureTitle.color, title.color);
            title.fontSize = DataViewObjects.getValue(objects, currentmeasurelabelprop.measureTitle.fontSize, title.fontSize);

            return title;
        }

        private getLabelSettings(dataView: DataView): ILabelSettings {
            let objects: DataViewObjects = null;
            let labelSettings: ILabelSettings;
            labelSettings = this.getDefaultLabelSettings();
            if (!dataView.metadata || !dataView.metadata.objects) {
                return labelSettings;
            }
            objects = dataView.metadata.objects;
            // tslint:disable-next-line:typedef
            const labelProps = props;
            labelSettings.color = DataViewObjects.getFillColor(objects, labelProps.labelSettings.color, labelSettings.color);
            labelSettings.fontSize = DataViewObjects.getValue(objects, labelProps.labelSettings.fontSize, labelSettings.fontSize);
            labelSettings.fontSize = labelSettings.fontSize > 25 ? 25 : labelSettings.fontSize;
            labelSettings.displayUnits = DataViewObjects.getValue(objects,
                                                                  labelProps.labelSettings.displayUnits, labelSettings.displayUnits);
            labelSettings.strokeWidth = DataViewObjects.getValue(objects,
                                                                 labelProps.labelSettings.strokeWidth, labelSettings.strokeWidth);
            if (labelSettings.strokeWidth > 4) {
                labelSettings.strokeWidth = 4;
            }

            return labelSettings;
        }

        public getAnimationSettings(dataView: DataView): IAnimationSettings {
            let objects: DataViewObjects = null;
            let settings: IAnimationSettings;
            settings = this.getDefaultAnimationSettings();
            if (!dataView.metadata || !dataView.metadata.objects) {
                return settings;
            }
            objects = dataView.metadata.objects;
            // tslint:disable-next-line:typedef
            let properties;
            properties = props;
            settings.show = DataViewObjects.getValue(objects, properties.animationSettings.show, settings.show);
            settings.duration = DataViewObjects.getValue(objects, properties.animationSettings.duration, settings.duration);
            settings.duration = settings.duration < 2 ? 2 : settings.duration > 20 ? 20 : settings.duration;

            return settings;
        }

        public enumerateObjectInstances(options: EnumerateVisualObjectInstancesOptions): VisualObjectInstanceEnumeration {
            // tslint:disable-next-line:typedef
            let objectName;
            objectName = options.objectName;
            let objectEnumeration: VisualObjectInstance[];
            objectEnumeration = [];
            let animationSettings: IAnimationSettings;
            animationSettings = this.getAnimationSettings(this.dataviews);
            let measureTitle: IMeasureTitle;
            measureTitle = this.getMeasureTitle(this.dataviews);
            let labels: ILabelSettings;
            labels = this.getLabelSettings(this.dataviews);

            switch (objectName) {
                case 'animationSettings':
                    objectEnumeration.push({
                        objectName: objectName,
                        displayName: 'Delay (seconds)',
                        selector: null,
                        properties: {
                            show: animationSettings.show,
                            duration: animationSettings.duration
                        }
                    });
                    break;
                case 'labelSettings':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            color: labels.color,
                            fontSize: labels.fontSize,
                            displayUnits: labels.displayUnits,
                            strokeWidth: labels.strokeWidth
                        },
                        selector: null
                    });
                    break;
                case 'measureTitle':
                    objectEnumeration.push({
                        objectName: objectName,
                        properties: {
                            color: measureTitle.color,
                            fontSize: measureTitle.fontSize
                        },
                        selector: null
                    });
                    break;
                case 'colorSelector':
                    let barDataPoint: IBarChartDataPoint;
                    for (barDataPoint of this.barDataPoints) {
                        objectEnumeration.push({
                            objectName: objectName,
                            displayName: barDataPoint.category,
                            properties: {
                                fill: {
                                    solid: {
                                        color: barDataPoint.color
                                    }
                                }
                            },
                            selector: barDataPoint.selectionId.getSelector()

                        });
                    }
                    break;
                default:
                    break;
            }

            return objectEnumeration;
        }
        // tslint:disable-next-line:no-any
        private static getTooltipData(value: any): VisualTooltipDataItem[] {
            return [{
                displayName: value.category,
                value: value.value.toString(),
                color: value.color
            }];
        }
    }
}
