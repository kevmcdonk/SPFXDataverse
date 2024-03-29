import * as React from 'react';

import 'jest';
// import { createRenderer } from 'react-test-renderer/shallow';
import { shallow } from 'enzyme';
import toJson from 'enzyme-to-json';

import SpfxDataverse from "./SpfxDataverse";

test('should render Sample component correctly', () => {
  /*
   * using the OOTB 'react-test-renderer'
   */
  // const renderer = createRenderer();
  // renderer.render(<Sample />);
  // console.log(renderer.getRenderOutput());
  // expect(renderer.getRenderOutput()).toMatchSnapshot();

  /*
   * using enzyme
   */
  const wrapper = shallow(<SpfxDataverse description="" isDarkTheme environmentMessage="" hasTeamsContext userDisplayName="" />);
  expect(wrapper.find('span').text()).toBe('Hello world:');
  expect(wrapper.find('li').length).toBe(3);
  // expect(toJson(wrapper)).toMatchSnapshot();
  expect(wrapper).toMatchSnapshot();
});