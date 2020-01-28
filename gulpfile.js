var gulp = require('gulp');
var del = require('del');

var ts = require('gulp-typescript');
var tsProject = ts.createProject('tsconfig.json');

gulp.task('clean', function() {
    // You can use multiple globbing patterns as you would with `gulp.src`
    return del(['dist']);
});

gulp.task('build', gulp.series(['clean'], function () {
    return tsProject.src()
        .pipe(tsProject())
        .js.pipe(gulp.dest('dist'));
}));

gulp.task('default', gulp.series(['clean', 'build']));

