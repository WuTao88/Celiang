import sys
import os
import math


def source_path(relative_path):
    if os.path.exists(relative_path):
        base_path = os.path.abspath(".")
    elif getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    return os.path.join(base_path, relative_path)





class CeLiang:

    def __init__(self,res='res\\RMC', zh0=0.0, x0=2921709.275, y0=550115.359,a=167, zhe=10581.786, xe=2917071.57167906,
                 ye=550689.5888589,ke=2.57254501945588):
        # jd [zh,X,Y,k,α,LS,R,Ls]
        # tpq [p1,q1,β1,T1,L,β2,p2,q2,T2]
        self.res=res
        self.bp = [zh0, x0, y0, 0, a, 0, 0, 0]
        self.ep = [zhe, xe, ye,ke]
    
    def SET(self,zh,rw):
        self.zh = zh
        self.RWD=rw
        self.JD1 = self.QXCS(zh)[0]
        self.JD2 = self.QXCS(zh)[1]

        self.TPQ2 = self.curvecs()
        self.widen = self.Roadcs()[1]
        self.highway = self.Roadcs()[0]

    '''
    直线函数
    '''

    def Line(self, zh, p):
        L0 = round(self.JD2[0] - self.JD1[0], 4)
        L = self.JD2[0] - zh
        k = self.JD2[3]
        x1 = self.JD2[1]
        y1 = self.JD2[2]
        f = math.pi / 2
        Y = round(y1 - L * math.sin(k) + p * math.sin(k + f), 4)
        X = round(x1 - L * math.cos(k) + p * math.cos(k + f), 4)
        return (X, Y)

    '''
    圆曲线函数
    '''

    def Circle(self, zh, p, JD, TP):
        R = JD[6]
        T1 = TP[3]
        ZH = self.Line(JD[0] - T1, 0)
        x = ZH[0]
        y = ZH[1]
        Ls1 = JD[5]
        L = zh - (JD[0] - T1 + Ls1)
        θ = L / R + TP[2]
        k = JD[3]
        f = -math.pi / 2 if JD[4] < 0 else math.pi / 2
        ki = k - θ if JD[4] < 0 else k + θ
        X0 = R * math.sin(θ) + TP[1]
        Y0 = R * (1 - math.cos(θ)) + TP[0]
        X = X0 * math.cos(k) + Y0 * math.cos(k + f) + x + p * math.cos(ki + math.pi / 2)
        Y = X0 * math.sin(k) + Y0 * math.sin(k + f) + y + p * math.sin(ki + math.pi / 2)

        return (X, Y)

    '''
    第一缓和曲线函数
    [p1,q1,β1,T1,L,β2,p2,q2,T2]
    [zh,X,Y,k,α,Ls1,R,Ls2]
    '''

    def FirstFlatCurve(self, zh, p, JD, TP):
        R = JD[6]
        zh1 = JD[0] - TP[3]
        L = zh - zh1
        L0 = JD[5]
        k = JD[3]
        θ = JD[4]
        ZH = self.Line(zh1, 0)
        x1 = ZH[0]
        y1 = ZH[1]
        f = -math.pi / 2 if θ < 0 else math.pi / 2
        θt = L ** 2 / (2 * R * L0)
        ki = k - θt if θ < 0 else k + θt
        # print("切线角：%s"%ki)
        X1 = L - L ** 5 / (40 * R ** 2 * L0 ** 2)
        Y1 = (L ** 3) / (6 * R * L0) - (L ** 7) / (336 * (R ** 3) * L0 ** 3)
        X = X1 * math.cos(k) + Y1 * math.cos(k + f) + x1 + p * math.cos(ki + math.pi / 2)
        Y = X1 * math.sin(k) + Y1 * math.sin(k + f) + y1 + p * math.sin(ki + math.pi / 2)

        return (X, Y)

    '''
    第二缓和曲线函数
    [p1,q1,β1,T1,L,β2,p2,q2,T2]
    [zh,X,Y,k,α,Ls1,R,Ls2]
    '''

    def SecondFlatCurve(self, zh, p, JD, TP):
        zh2 = JD[0] - TP[3] + JD[5] + JD[7] + TP[4]

        L0 = JD[7]
        θ = JD[4]
        L = zh2 - zh
        R = JD[6]

        k1 = JD[3] + JD[4]
        x1 = JD[1] + math.cos(k1) * TP[8]
        y1 = JD[2] + math.sin(k1) * TP[8]
        k2 = JD[3] + JD[4] + math.pi
        # print((x1,y1))
        θt = L ** 2 / (2 * R * L0)
        ki = k2 + θt if θ < 0 else k2 - θt
        f = -math.pi / 2 if θ > 0 else math.pi / 2
        # print(L)
        X1 = L - L ** 5 / (40 * R ** 2 * L0 ** 2)
        Y1 = (L ** 3) / (6 * R * L0) - (L ** 7) / (336 * (R ** 3) * L0 ** 3)
        X = X1 * math.cos(k2) + Y1 * math.cos(k2 + f) + x1 + p * math.cos(ki + math.pi / 2)
        Y = X1 * math.sin(k2) + Y1 * math.sin(k2 + f) + y1 + p * math.sin(ki + math.pi / 2)
        return (X, Y)

    '''
    直坡
    '''

    def LineSlope(self, x1=0.0, h1=0.0, k=0.0, zh=0.0, jd1=0.0, jd2=0.0):
        L0 = jd2 - jd1
        L = zh - jd1
        Y = round(h1 + L * math.sin(k), 4)
        X = round(x1 + L * math.cos(k), 4)
        return (X, Y)

    '''
    缓坡
    '''

    def CircleSlope(self, x1, y1, β1, β2, R, zh, jd):
        T0 = R * math.tan((β1 - β2) / 2)
        X1 = x1 - T0 * math.cos(β1)
        Y1 = y1 - T0 * math.sin(β1)
        L = zh - (jd - T0)
        θ = L / R
        C = 2 * R * math.sin(θ / 2)
        if β1 < β2:
            X = round(X1 + C * math.cos(β1 + θ / 2), 4)
            Y = round(Y1 + C * math.sin(β1 + θ / 2), 4)
        else:
            X = round(X1 + C * math.cos(β1 - θ / 2), 4)
            Y = round(Y1 + C * math.sin(β1 - θ / 2), 4)
        return (X, Y)

    '''
    高程
    '''

    def Height(self, pl=0.0, filename=None):
        K = self.zh

        name = self.res + '\\sqx' if filename == None else filename
        data = open(name, "r+", encoding="UTF-8")
        cs = []
        gc = {}
        p = (0, 2079.647)
        β1 = 0
        for d in data:
            gc = eval(d)
        KEYS = list(gc.keys())
        for key in KEYS:
            if KEYS.index(key) > 0:
                x1 = KEYS[KEYS.index(key) - 1]
                x2 = key
                y1 = gc[x1][0]
                y2 = gc[x2][0]
                r1 = gc[x1][1]
                r2 = gc[x2][1]
                β1 = gc[x1][2]
                β2 = gc[x2][2]
                p = self.LineSlope(p[0], p[1], β1, x2, x1, x2)
                L = r2 * (β1 - β2)
                T = r2 * math.tan((β1 - β2) / 2)
                k1 = round(x2 - T, 4)
                k2 = round(k1 + L, 4)
                if K <= k1:
                    p = self.LineSlope(p[0], p[1], β1, x1 - x2 + K, x1, x2)

                    

                    H = round(p[1]  + abs(pl) * self.highway[0], 3) if pl < 0 else round(
                            p[1] + abs(pl) * self.highway[1], 3)
                    return (K, H)
                    break
                elif K > k1 and K <= k2:
                    p = self.CircleSlope(p[0], p[1], β1, β2, r2, K, x2)


                    H = round(p[1] + abs(pl) * self.highway[0], 3) if pl < 0 else round(
                            p[1] + abs(pl) * self.highway[1], 3)
                    return (K, H)
                    break

    '''
    切线及圆曲线长度
    
    '''

    def curvecs(self):

        # =========================
        if self.JD2 == self.ep :
            cs2 = [0, 0, 0, 0, 0, 0, 0, 0, 0]
        else:
            α = abs(self.JD2[4])
            Ls1 = self.JD2[5]
            R = self.JD2[6]
            Ls2 = self.JD2[7]
            q1 = round(Ls1 / 2 - (Ls1 ** 3) / (240 * (R ** 2)) + (Ls1 ** 6) / (34560 * (R ** 4)), 4)
            p1 = round((Ls1 ** 2) / (24 * R) - (Ls1 ** 4) / (2688 * (R ** 3)), 4)
            L = round(α * R - (Ls1 + Ls2) / 2, 4)
            q2 = round(Ls2 / 2 - (Ls2 ** 3) / (240 * (R ** 2)) + (Ls2 ** 6) / (34560 * (R ** 4)), 4)
            p2 = round((Ls2 ** 2) / (24 * R) - (Ls2 ** 4) / (2688 * (R ** 3)), 4)
            β1 = round(Ls1 / (2 * R), 6)
            β2 = round(Ls2 / (2 * R), 6)
            T2 = round((p1 - p2) / (2 * math.tan(α / 2)) + 0.5 * (p1 + p2 + 2 * R) * (math.tan(α / 2)) + q2, 4)
            T1 = round((p2 - p1) / (2 * math.tan(α / 2)) + 0.5 * (p1 + p2 + 2 * R) * (math.tan(α / 2)) + q1, 4)
            cs2 = [p1, q1, β1, T1, L, β2, p2, q2, T2]
        return cs2

    '''
    平曲线参数
    '''

    def QXCS(self, zh=0.0, filename=None):

        name = self.res + "\\pqx" if filename == None else filename
        # JD 坐标x,y

        PQX = {}
        # {桩号KN:[X,Y,k,α,Ls1,R,Ls2]}
        cs = open(name, "r+", encoding="UTF-8")

        for i in cs:
            PQX = eval(i)
        JDKS = list(PQX.keys())
        for KN in JDKS:
            α = abs(PQX[KN][3])
            Ls1 = PQX[KN][4]
            R = PQX[KN][5]
            Ls2 = PQX[KN][6]
            q1 = round(Ls1 / 2 - (Ls1 ** 3) / (240 * (R ** 2)) + (Ls1 ** 6) / (34560 * (R ** 4)), 4)
            q2 = round(Ls2 / 2 - (Ls2 ** 3) / (240 * (R ** 2)) + (Ls2 ** 6) / (34560 * (R ** 4)), 4)
            p1 = round((Ls1 ** 2) / (24 * R) - (Ls1 ** 4) / (2688 * (R ** 3)), 4)
            L = round(α * R - (Ls1 + Ls2) / 2, 4)

            p2 = round((Ls2 ** 2) / (24 * R) - (Ls2 ** 4) / (2688 * (R ** 3)), 4)
            β1 = round(Ls1 / (2 * R), 6)
            β2 = round(Ls2 / (2 * R), 6)
            T2 = round((p1 - p2) / (2 * math.tan(α / 2)) + 0.5 * (p1 + p2 + 2 * R) * math.tan(α / 2) + q2, 4)
            T1 = round((p2 - p1) / (2 * math.tan(α / 2)) + 0.5 * (p1 + p2 + 2 * R) * math.tan(α / 2) + q1, 4)

            if zh <= KN - T1 + Ls1 + L + Ls2:

                jd1 = None
                if JDKS.index(KN) == 0:
                    jd1 = self.bp

                else:
                    jd1 = PQX[JDKS[JDKS.index(KN) - 1]]
                    jd1.insert(0, JDKS[JDKS.index(KN) - 1])

                jd2 = PQX[KN]
                jd2.insert(0, KN)
                return (jd1, jd2)
            elif zh>10547.874:
                jd1=PQX[JDKS[-1]]
                jd1.insert(0, JDKS[- 1])
                jd2=self.ep
                return (jd1, jd2)

    '''
    加宽
    '''

    def WIDEN(self, r=0.0):
        wides = open(self.res + '\\' + 'jiakuan', 'r+', encoding='UTF-8')
        JiaKuan = 0
        xb = []
        for b in wides:
            xb = eval(b)
        for b in xb:
            if r >= b[0] and r < b[1]:
                JiaKuan = b[2]

        return JiaKuan

    '''
    超高
    
    '''

    def HIGHWAY(self, r=0.0):
        exHc = open(self.res + '\\' + 'chaogao', 'r+', encoding='UTF-8')

        ChaoGao = (-0.02,-0.02)
        slops = []
        for hc in exHc:
            slops = eval(hc)
        for hc in slops:
            if r >= hc[0] and r < hc[1]:
                ChaoGao = (hc[2],-hc[2])
        return ChaoGao

    '''
    坐标函数
    '''

    def Point(self, p=0.0):
        zh = self.zh

        P = None

        if zh <= self.JD2[0] - self.TPQ2[3]:
            # print('直线')
            P = self.Line(zh, p)

        elif self.JD2[5] != 0 and zh <= self.JD2[0] - self.TPQ2[3] + self.JD2[5]:

            # print('第一缓和曲线')
            P = self.FirstFlatCurve(zh, p, self.JD2, self.TPQ2)


        elif zh <= self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4]:

            # print('圆曲线')
            P = self.Circle(zh, p, self.JD2, self.TPQ2)


        elif self.JD2[7] != 0 and zh <= self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4] + self.JD2[7]:

            # print('第二缓和曲线')
            P = self.SecondFlatCurve(zh, p, self.JD2, self.TPQ2)

        return P

    '''
    路面宽度，横坡
    '''

    def Roadcs(self):
        zh = self.zh
        α = self.JD2[4] if self.JD2!=self.ep else 0
        R = self.JD2[6] if self.JD2!=self.ep else 0
        widen = self.WIDEN(R)
        highway = self.HIGHWAY(R)
        Roadi = None
        Roadw = None
        RDHF=self.RWD/2
        if zh <= self.JD2[0] - self.TPQ2[3]:
            #print('直线')
            Roadw = (-RDHF, RDHF)
            Roadi = (-0.02, -0.02)
        elif self.JD2[5] != 0 and zh <= self.JD2[0] - self.TPQ2[3] + self.JD2[5]:
            #print('第一缓曲')
            if α > 0:
                Roadw = (-RDHF, round(RDHF + (zh - (self.JD2[0] - self.TPQ2[3])) * widen / self.JD2[5], 2))
                Roadi = (round(-0.02 + (zh - (self.JD2[0] - self.TPQ2[3])) * (highway[0] + 0.02) / self.JD2[5], 3),
                         round(-0.02 + (zh - (self.JD2[0] - self.TPQ2[3])) * (highway[1] + 0.02) / self.JD2[5], 3))
            else:
                Roadw = (round(-RDHF - (zh - (self.JD2[0] - self.TPQ2[3])) * widen / self.JD2[5], 2), RDHF)
                Roadi = (round(-0.02 + (zh - (self.JD2[0] - self.TPQ2[3])) * (-highway[0] + 0.02) / self.JD2[5], 3),
                         round(-0.02 + (zh - (self.JD2[0] - self.TPQ2[3])) * (-highway[1] + 0.02) / self.JD2[5], 3))
            Roadi = (highway[0],highway[1]) if highway[0]==highway[1] else Roadi
        elif zh <= self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4]:
            #print('圆')
            if α > 0:
                Roadw = (-RDHF, RDHF + widen)
                Roadi = (highway[0], highway[1])
            else:
                Roadw = (-RDHF - widen, RDHF)
                Roadi = (-highway[0], -highway[1])  
            Roadi = (highway[0],highway[1]) if highway[0]==highway[1] else Roadi
        elif self.JD2[7] != 0 and zh <= self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4] + self.JD2[7]:
            #print('第二缓曲')
            if α > 0:
                Roadw = (-RDHF, round(
                    RDHF + (zh - (self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4])) * widen / self.JD2[7], 2))
                Roadi = (round(
                    -0.02 + (zh - (self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4])) * (highway[0] + 0.02) /
                    self.JD2[7], 3), round(
                    -0.02 + (zh - (self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4])) * (highway[1] + 0.02) /
                    self.JD2[7], 3))
            else:
                Roadw = (
                round(-RDHF - (zh - (self.JD2[0] - self.TPQ2[3] + self.JD2[5] + self.TPQ2[4])) * widen / self.JD2[7],
                      2), RDHF)
                Roadi = (round(-0.02 + (zh - (self.JD2[0] - self.TPQ2[3])) * (-highway[0] + 0.02) / self.JD2[7], 3),
                         round(-0.02 + (zh - (self.JD2[0] - self.TPQ2[3])) * (-highway[1] + 0.02) / self.JD2[7], 3))
            Roadi = (highway[0],highway[1]) if highway[0]==highway[1] else Roadi             
        return [Roadi, Roadw]

    def side(self, position="左侧"):
        if position == '左侧':
            p = self.widen[0]
            P = self.Point(p)
            GC = self.Height(p)
        else:
            p = self.widen[1]
            P = self.Point(p)
            GC = self.Height(p)

        return [GC[0], p, GC[1], P]
        
