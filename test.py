T = int(input())
for i in range(T):
    x1, y1, r1, x2, y2, r2 = map(int, input().split())
    dist = (x1-x2)**2 + (y1-y2)**2
    dist2 = (r1+r2)**2
    sameYn = (x1 == x2 and y1 == y2)
    if sameYn:
        if r1 == r2:
            print(-1)
        else:
            print(0)
    elif dist < dist2:
        print(2)
    elif dist == dist2:
        print(1)
    else:
        print(0)