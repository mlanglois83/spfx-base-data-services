var gulp = require('gulp');
var del = require('del');

var ts = require('gulp-typescript');
var tsProject = ts.createProject('tsconfig.json');
var sourcemaps = require('gulp-sourcemaps');
var eslint = require('gulp-eslint');
var gulpIf = require('gulp-if');


gulp.task('eslint', () => {
    const hasFixFlag = process.argv.slice(2).includes('--fix');
    return tsProject.src()
        // eslint() attaches the lint output to the "eslint" property
        // of the file object so it can be used by other modules.
        .pipe(eslint({fix: hasFixFlag}))
        // eslint.format() outputs the lint results to the console.
        // Alternatively use eslint.formatEach() (see Docs).
        .pipe(eslint.format())
});

gulp.task('clean', function() {
    // You can use multiple globbing patterns as you would with `gulp.src`
    return del(['dist']);
});
gulp.task('build', function() {
    return tsProject.src()
        .pipe(sourcemaps.init())
        .pipe(tsProject())
        .pipe(sourcemaps.write('.', {includeContent: false, sourceRoot:"."}))
        .pipe(gulp.dest('dist'));
});

gulp.task('default', gulp.series(['clean', "eslint", 'build']));

