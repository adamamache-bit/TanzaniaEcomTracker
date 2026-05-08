export const DEFAULT_EXCHANGE_RATE = 2850;

function toNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

export function resolveExchangeRate(exchangeRate = DEFAULT_EXCHANGE_RATE) {
  const parsed = toNumber(exchangeRate);
  return parsed > 0 ? parsed : DEFAULT_EXCHANGE_RATE;
}

export function tshToUsd(amountTsh, exchangeRate = DEFAULT_EXCHANGE_RATE) {
  return toNumber(amountTsh) / resolveExchangeRate(exchangeRate);
}

export function usdToTsh(amountUsd, exchangeRate = DEFAULT_EXCHANGE_RATE) {
  return toNumber(amountUsd) * resolveExchangeRate(exchangeRate);
}

export function formatMoneyBoth(amountUsd, exchangeRate = DEFAULT_EXCHANGE_RATE) {
  const usd = toNumber(amountUsd);
  const tsh = usdToTsh(usd, exchangeRate);
  return {
    amountUsd: usd,
    amountTsh: tsh,
    label: `$${usd.toFixed(2)} | TSh ${Math.round(tsh).toLocaleString()}`,
  };
}
