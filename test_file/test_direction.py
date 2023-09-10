def direction_switcher(angle):
    if (angle == 0):
        diro = 'N'
    elif (angle < 45):
        diro = 'NNE'
    elif (angle == 45):
        diro = 'NE'
    elif (45 < angle < 90):
        diro = 'ENE'
    elif (angle == 90):
        diro = 'E'
    elif (90 < angle < 135):
        diro = 'ESE'
    elif (angle == 135):
        diro = 'SE'
    elif (135 < angle < 180):
        diro = 'SSE'
    elif (angle == 180):
        diro = 'S'
    elif (180 < angle < 225):
        diro = 'SSW'
    elif (angle == 225):
        diro = 'SW'
    elif (225 < angle < 270):
        diro = 'SSW'
    elif (angle == 270):
        diro = 'W'
    elif (270 < angle < 315):
        diro = 'WNW'
    elif (angle == 315):
        diro = 'NW'
    elif (315 < angle < 360):
        diro = 'NNW'
    return (diro)

print (direction_switcher(25))
