[data,R_t,A_t,Am_t,A_arr,Am_arr,R_arr] = composite_excel()
function [data,R_t,A_t,Am_t,A_arr,Am_arr,R_arr] = composite_excel()
A_t = 0;
Am_t = 0;
R_t = 0;
data = readtable('CurvedBeams.xlsx');
type = data.type;
xy = size(data);
N = xy(1);
A_arr = zeros(N,1);
Am_arr = zeros(N,1);
R_arr = zeros(N,1);
for i = 1:N
    ty = cell2mat(type(i));
    if ty == "r"
        a = data.a(i);
        b = data.b(i);
        c = data.c(i);
        [A,Am,R] = Rect(a,b,c);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "t"
        a = data.a(i);
        b = data.b(i);
        c = data.c(i);       
        [A,Am,R] = Triangle(a,b,c);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "tr"
        a = data.a(i);
        b1 = data.b1(i);
        b2 = data.b2(i);
        c = data.c(i);
        [A,Am,R] = Trapezium(a,b1,b2,c);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "c"
        R = data.R(i);
        b = data.b(i);
        [A,Am,R] = Circle(R,b);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "e"
        h = data.h(i);
        b = data.b(i);
        R = data.R(i);
        [A,Am,R] = Ellipse(h,b,R);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "cc"
        R  = data.R(i);
        b1 = data.b1(i);
        b2 = data.b2(i);
        [A,Am,R] = ConcentricCircle(R,b1,b2);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "ce"
        R = data.R(i);
        b1 = data.b1(i);
        b2 = data.b2(i);
        h1 = data.h1(i);
        h2 = data.h2(i);
        [A,Am,R] = ConcentricEllipse(R,b1,b2,h1,h2);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "an"
        a = data.a(i);
        b = data.b(i);
        theta = data.theta(i);
        [A,Am,R] = near_apex_arc(a,b,theta);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    elseif ty == "aa"
        a = data.a(i);
        b = data.b(i);
        theta =  data.theta(i);
        [A,Am,R] = away_apex_arc(a,b,theta);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    else
        a = data.a(i);
        b = data.b(i);
        h = data.h(i);
        [A,Am,R] = away_apex_arc_dim(a,b,h);
        A_arr(i) = A;
        Am_arr(i) = Am;
        R_arr(i) = R;
        A_t = A_t + A;
        Am_t = Am_t + Am;
        R_t = R_t + R*A;
    end
end
R_t = R_t/A_t;
data.A = A_arr;
data.Am = Am_arr;
data.R = R_arr;
writetable(data,'CurvedBeamsSolution.xlsx');
end

function [A,Am,R] = Rect(a,b,c)
A = b*(c - a);
Am = b*log(c/a);
R = (a + c)/2;
end

function [A,Am,R] = Triangle(a,b,c)
A = b/2*(c - a);
Am = (b*c)/(c-a)*log(c/a) - b;
R = (2*a + c)/3;
end

function [A,Am,R] = Trapezium(a,b1,b2,c)
A = (b1 + b2)/2*(c - a);
Am = (b1*c - b2*a)/(c-a)*log(c/a) - b1 +b2 ; 
R = (a*(2*b1 + b2) + c*(b1 + 2*b2))/(3*(b1 + b2));
end

function [A,Am,R] = Circle(R,b)
A = pi*b^2;
Am = 2*pi*(R - sqrt(R^2 - b^2));
end

function [A,Am,R] = Ellipse(h,b,R)
A = pi*b*h;
Am = 2*pi*b/h*(R - sqrt(R^2 - h^2));
end

function [A,Am,R] = ConcentricCircle(R,b1,b2)
A = pi*(b1^2 - b2^2);
Am = 2*pi*(sqrt(R^2 - b2^2) - sqrt(R^2 - b1^2));
end

function [A,Am,R] = ConcentricEllipse(R,b1,b2,h1,h2)
A = pi*(b1*h1 - b2*h2);
Am = 2*pi*(b1*R/h1 - b2*R/h2 - b1/h1*sqrt(R^2 - h1^2) + b2/h2*(R^2 - h2^2));
end

function [A,Am,R] = near_apex_arc(a,b,theta)
A = b^2*theta - b^2/2*sin(2*theta);
R = a + (4*b*sin(theta)^3)/(3*(2*theta - sin(2*theta)));
if a>=b
    Am = 2*a*theta - 2*b*sin(theta) - pi*sqrt(a^2 - b^2) + 2*sqrt(a^2 - b^2)*asin((b + a*cos(theta))/(a + b*cos(theta)));
else
    Am = 2*a*theta - 2*b*sin(theta) + 2*sqrt(b^2 - a^2)*log((b + a*cos(theta) + sqrt(b^2 - a^2)*sin(theta))/(a + b*cos(theta)));
end
end

function [A,Am,R] = away_apex_arc(a,b,theta)
A = b^2*theta - b^2/2*sin(2*theta);
R = a + (4*b*sin(theta)^3)/(3*(2*theta - sin(2*theta)));
Am = 2*a*theta + 2*b*sin(theta) - pi*sqrt(a^2 - b^2) - 2*sqrt(a^2 - b^2)*asin((b - a*cos(theta))/(a - b*cos(theta)));
end

function [A,Am,R] = away_apex_arc_dim(a,b,h)
A = pi*b*h/2;
R = a - 4*h/(3*pi);
Am = 2*b + pi*b/h*(a - sqrt(a^2 - h^2)) - 2*b/h*sqrt(a^2 - h^2)*asin(h/a);
end
