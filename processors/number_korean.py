# 숫자 → 한글 금액 변환

_ONES = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구']
_UNITS = ['', '십', '백', '천']
_BIG   = ['', '만', '억', '조']

def number_to_korean(n: int) -> str:
    """375577990 → '삼억칠천오백오십칠만칠천구백구십'"""
    if n == 0:
        return '영'
    if n < 0:
        return '마이너스 ' + number_to_korean(-n)

    result = ''
    group_idx = 0

    while n > 0:
        group = n % 10000
        n //= 10000

        if group != 0:
            group_str = ''
            for unit_idx in range(3, -1, -1):
                digit = (group // (10 ** unit_idx)) % 10
                if digit == 0:
                    continue
                # '일십', '일백', '일천'은 '십', '백', '천'으로 (단, 일의 자리 '일'은 표기)
                if digit == 1 and unit_idx > 0:
                    group_str += _UNITS[unit_idx]
                else:
                    group_str += _ONES[digit] + _UNITS[unit_idx]
            result = group_str + _BIG[group_idx] + result

        group_idx += 1

    return result


def amount_to_korean_formal(n: int) -> str:
    """영수증용: 375577990 → '삼억칠천오백오십칠만칠천구백구십'"""
    return number_to_korean(n)


if __name__ == '__main__':
    tests = [375_577_990, 13_125_000, 46_000_000, 316_452_990, 1000, 10000, 100000000]
    for t in tests:
        print(f"{t:,} → {number_to_korean(t)}")
