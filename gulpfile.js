var gulp = require('gulp');
var del = require('del');
var spawn = require('child_process').spawn;

var ts = require('gulp-typescript');
var tsProject = ts.createProject('tsconfig.json');

gulp.task('clean', function() {
    // You can use multiple globbing patterns as you would with `gulp.src`
    return del(['dist']);
});

gulp.task('default', function () {
    return tsProject.src()
        .pipe(tsProject())
        .js.pipe(gulp.dest('dist'));
});


gulp.task('publish', function (done) {
    pawn('npm', ['publish'], { stdio: 'inherit' }).on('close', done);
});
gulp.task('version-patch', function (done) {
    pawn('npm', ['version', 'patch'], { stdio: 'inherit' }).on('close', done);
});