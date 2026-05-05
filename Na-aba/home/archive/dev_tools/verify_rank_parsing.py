
import re
import pandas as pd

def test_rank_parsing():
    ranks = ["이사", "부장", "차장", "과장", "대리", "계장", "주임", "기사"]
    rank_pattern = re.compile(r"\[?(이사|부장|차장|과장|대리|계장|주임|기사)\]?\s*(.*)")
    
    worker_rank_map = {
        "주진철": "부장", "우명광": "대리", "김진환": "주임", "장승대": "계장",
        "김성렬": "주임", "박광복": "부장", "주영광": "과장"
    }
    
    test_cases = [
        ("[부장] 주진철", "부장", "주진철"),
        ("부장 주진철", "부장", "주진철"),
        ("주진철", "부장", "주진철"),
        ("과장 주영광", "과장", "주영광"),
        ("[이사] 홍길동", "이사", "홍길동"),
        ("대리 우명광", "대리", "우명광"),
        ("부장", "부장", "부장"), # Edge case: only rank
    ]
    
    print(f"{'Input':<20} | {'Expected Rank':<15} | {'Detected Rank':<15} | {'Match'}")
    print("-" * 65)
    
    all_passed = True
    for raw_worker, exp_rank, exp_name in test_cases:
        match = rank_pattern.search(raw_worker)
        rank = None
        worker_name_only = raw_worker
        
        if match:
            if match.group(1) in ranks:
                rank = match.group(1)
                worker_name_only = match.group(2).strip()
                if not worker_name_only:
                    worker_name_only = rank
        
        if not rank:
            clean_name = re.sub(r'\(.*?\)', '', raw_worker).strip()
            rank = worker_rank_map.get(clean_name)
            worker_name_only = clean_name
            
        is_match = (rank == exp_rank)
        print(f"{raw_worker:<20} | {exp_rank:<15} | {str(rank):<15} | {'OK' if is_match else 'FAIL'}")
        if not is_match:
            all_passed = False
            
    if all_passed:
        print("\nAll rank parsing tests passed!")
    else:
        print("\nSome rank parsing tests failed.")

if __name__ == "__main__":
    test_rank_parsing()
