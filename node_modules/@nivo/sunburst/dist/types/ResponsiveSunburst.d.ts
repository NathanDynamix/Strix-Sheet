import { ReactElement } from 'react';
import { ResponsiveProps, WithChartRef } from '@nivo/core';
import { SunburstSvgProps } from './types';
type ResponsiveSunburstProps<RawDatum> = ResponsiveProps<Partial<Omit<SunburstSvgProps<RawDatum>, 'data' | 'width' | 'height'>> & Pick<SunburstSvgProps<RawDatum>, 'data'>>;
export declare const ResponsiveSunburst: <RawDatum>(props: WithChartRef<ResponsiveSunburstProps<RawDatum>, SVGSVGElement>) => ReactElement;
export {};
//# sourceMappingURL=ResponsiveSunburst.d.ts.map