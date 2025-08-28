#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
사이즈 매칭 로직 테스트 스크립트
"""

from match import PerfectMatcher

def test_size_matching():
    matcher = PerfectMatcher()
    
    print("🔍 새로운 사이즈 매칭 로직 테스트")
    print("=" * 50)
    
    # 테스트 케이스들
    test_cases = [
        # (입력, 예상 출력)
        ("7(M)", {"7", "m"}),  # 괄호 안팎 모두 추출
        ("M(9)", {"m", "9"}),  # 괄호 안팎 모두 추출
        ("L", {"l"}),          # 괄호 없음
        ("9", {"9"}),          # 숫자만
        ("XL(100)", {"xl", "100"}),  # 사이즈 + 숫자
        ("7(M), 9(L)", {"7", "m", "9", "l"}),  # 여러 사이즈
    ]
    
    for input_size, expected in test_cases:
        result = matcher.extract_size_parts_ultra(input_size)
        print(f"입력: '{input_size}'")
        print(f"결과: {result}")
        print(f"예상: {expected}")
        print(f"매칭: {'✅' if result == expected else '❌'}")
        print("-" * 30)
    
    print("\n🎯 매칭 시나리오 테스트")
    print("=" * 50)
    
    # 실제 매칭 시나리오 테스트
    scenarios = [
        ("7", "7(M)", "✅ 매칭 성공 (괄호 밖 값 일치)"),
        ("M(9)", "9", "✅ 매칭 성공 (괄호 안 값 일치)"),
        ("7(M)", "L(9)", "❌ 매칭 실패 (둘 다 불일치)"),
        ("M", "M(9)", "✅ 매칭 성공 (괄호 밖 값 일치)"),
        ("XL", "XL(110)", "✅ 매칭 성공 (괄호 밖 값 일치)"),
    ]
    
    for size1, size2, expected_result in scenarios:
        set1 = matcher.extract_size_parts_ultra(size1)
        set2 = matcher.extract_size_parts_ultra(size2)
        
        # 교집합이 있으면 매칭 성공
        has_intersection = bool(set1.intersection(set2))
        
        print(f"'{size1}' vs '{size2}'")
        print(f"  {size1} → {set1}")
        print(f"  {size2} → {set2}")
        print(f"  교집합: {set1.intersection(set2)}")
        print(f"  결과: {'✅ 매칭' if has_intersection else '❌ 불일치'}")
        print(f"  예상: {expected_result}")
        print("-" * 40)

if __name__ == "__main__":
    test_size_matching() 