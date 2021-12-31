import { i2a } from './i2a'

function testi2a() {
  const results = [{ n: 1, s: 'A' }]
  results.forEach(i => {
    if (i2a(i.n) !== i.s) {
      console.error(i)
    }
  })
}
