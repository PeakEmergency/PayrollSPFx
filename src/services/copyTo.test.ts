import { copyProblem, recipients } from './copyTo';

it('adds the copy address to the To line', () => {
  expect(recipients('accounts@x.com', undefined)).toBe('accounts@x.com');
  expect(recipients('accounts@x.com', ' me@x.com ')).toBe('accounts@x.com;me@x.com');
});

it('checks the copy address only when a copy is wanted', () => {
  expect(copyProblem(undefined)).toBe('');
  expect(copyProblem('me@x.com')).toBe('');
  expect(copyProblem('')).not.toBe('');
  expect(copyProblem('me@x.com;evil@y.com')).not.toBe('');
  expect(copyProblem('not an email')).not.toBe('');
});
